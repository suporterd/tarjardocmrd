// Elements
const uploadScreen = document.getElementById('upload-screen');
const editorScreen = document.getElementById('editor-screen');
const loadingScreen = document.getElementById('loading-screen');
const loadingText = document.getElementById('loading-text');

const fileInput = document.getElementById('file-input');
const dropZone = document.getElementById('drop-zone');

const pdfContainer = document.getElementById('pdf-container');
const pdfCanvas = document.getElementById('pdf-canvas');
const overlayContainer = document.getElementById('overlay-container');

const btnClose = document.getElementById('btn-close');
const btnExport = document.getElementById('btn-export');
const btnPrev = document.getElementById('btn-prev');
const btnNext = document.getElementById('btn-next');
const pageInfo = document.getElementById('page-info');

// State
// Configuração do Painel de Auditoria (Power Automate)
// Cole a URL do seu gatilho HTTP do Power Automate abaixo:
const WEBHOOK_URL = 'https://default975003b969304fb3bbe11f2ec3c14c.d0.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/70e9f7582b5b48e0b1f3a8cf386bc292/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=veoCUlGOUWDCkl65LTdf9isrrosnqeC4eTU5osvVeUE'; 

let originalPdfBytes = null;
let currentFileName = ''; // Guardará o nome do arquivo original
let currentUserEmail = 'Desconhecido';
let currentUserName = 'Fora do Teams (Navegador)';

let pdfDoc = null; // pdf.js doc
let currentPageNum = 1;
let currentRenderTask = null;
let pdfScale = 1.5; // Render scale
let currentViewport = null;

// Tenta inicializar o SDK do Teams para capturar a identidade
if (typeof microsoftTeams !== 'undefined') {
    microsoftTeams.app.initialize().then(() => {
        microsoftTeams.app.getContext().then((context) => {
            if (context && context.user) {
                currentUserName = context.user.userPrincipalName || context.user.displayName || 'Usuário Teams';
                currentUserEmail = context.user.userPrincipalName || 'sem-email@aedas';
            }
        }).catch(e => console.warn("Não foi possível obter o contexto de usuário."));
    }).catch(e => console.warn("Rodando fora do Teams. Identidade anônima assumida."));
}

// redactions[pageNumber] = [ { x, y, width, height } ... relative to unscaled PDF points or viewport? ]
// Let's store them relative to the viewport at current scale, then convert during export, 
// OR better yet, store them in PDF point coordinates immediately.
// Actually, since we only show one page at a time, we can just store them as normalized coordinates (0 to 1) 
// or PDF points to easily re-render them if we change scale.
// Let's store them as { x, y, w, h } in DOM pixels at the current scale, but we will clear and recreate DOM elements on page change.
// To persist across pages, we need to store them in a state object.
let redactions = {}; // { pageNum: [ {x, y, width, height, id} ] } // stored in PDF Points (unscaled, from top-left)

let isDrawing = false;
let startX, startY;
let selectionBox = null;

// Initialize
function showScreen(screen) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    screen.classList.add('active');
}

function showLoading(text) {
    loadingText.textContent = text;
    showScreen(loadingScreen);
}

// Upload Handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
});

btnClose.addEventListener('click', () => {
    showScreen(uploadScreen);
    originalPdfBytes = null;
    pdfDoc = null;
    redactions = {};
    fileInput.value = '';
});

async function handleFile(file) {
    if (file.type !== 'application/pdf') {
        alert('Por favor, selecione um arquivo PDF.');
        return;
    }

    showLoading('Carregando documento...');

    try {
        currentFileName = file.name;
        const arrayBuffer = await file.arrayBuffer();
        originalPdfBytes = new Uint8Array(arrayBuffer);

        // Load with PDF.js (pass a slice to avoid the worker detaching the buffer)
        const loadingTask = pdfjsLib.getDocument({ data: originalPdfBytes.slice() });
        pdfDoc = await loadingTask.promise;
        
        currentPageNum = 1;
        redactions = {};
        
        showScreen(editorScreen);
        renderPage(currentPageNum);
    } catch (error) {
        console.error(error);
        alert('Erro ao carregar o PDF. O arquivo pode estar corrompido ou protegido com senha.');
        showScreen(uploadScreen);
    }
}

// PDF.js Rendering
async function renderPage(pageNum) {
    if (currentRenderTask) {
        currentRenderTask.cancel();
    }

    // Update pagination UI
    pageInfo.textContent = `Pág. ${pageNum} de ${pdfDoc.numPages}`;
    btnPrev.disabled = pageNum <= 1;
    btnNext.disabled = pageNum >= pdfDoc.numPages;

    try {
        const page = await pdfDoc.getPage(pageNum);
        
        // Calculate scale to fit screen width (optional, hardcoded for now)
        // A good default scale for readability on web is 1.5
        const viewport = page.getViewport({ scale: pdfScale });
        currentViewport = viewport;

        const ctx = pdfCanvas.getContext('2d');
        pdfCanvas.height = viewport.height;
        pdfCanvas.width = viewport.width;
        
        // Match container size to canvas
        pdfContainer.style.width = `${viewport.width}px`;
        pdfContainer.style.height = `${viewport.height}px`;

        const renderContext = {
            canvasContext: ctx,
            viewport: viewport
        };

        currentRenderTask = page.render(renderContext);
        await currentRenderTask.promise;
        
        // Render existing redactions for this page
        renderRedactionElements();
        
    } catch (error) {
        if (error.name !== 'RenderingCancelledException') {
            console.error('Error rendering page:', error);
        }
    }
}

btnPrev.addEventListener('click', () => {
    if (currentPageNum <= 1) return;
    currentPageNum--;
    renderPage(currentPageNum);
});

btnNext.addEventListener('click', () => {
    if (currentPageNum >= pdfDoc.numPages) return;
    currentPageNum++;
    renderPage(currentPageNum);
});

// Drawing logic
overlayContainer.addEventListener('mousedown', (e) => {
    if (e.target.classList.contains('delete-btn')) return;
    
    isDrawing = true;
    const rect = overlayContainer.getBoundingClientRect();
    startX = e.clientX - rect.left;
    startY = e.clientY - rect.top;

    if (!selectionBox) {
        selectionBox = document.createElement('div');
        selectionBox.id = 'selection-box';
        overlayContainer.appendChild(selectionBox);
    }
    
    selectionBox.style.left = `${startX}px`;
    selectionBox.style.top = `${startY}px`;
    selectionBox.style.width = '0px';
    selectionBox.style.height = '0px';
    selectionBox.style.display = 'block';
});

overlayContainer.addEventListener('mousemove', (e) => {
    if (!isDrawing) return;
    
    const rect = overlayContainer.getBoundingClientRect();
    let currentX = e.clientX - rect.left;
    let currentY = e.clientY - rect.top;
    
    // Clamp to boundaries
    currentX = Math.max(0, Math.min(currentX, rect.width));
    currentY = Math.max(0, Math.min(currentY, rect.height));

    const width = Math.abs(currentX - startX);
    const height = Math.abs(currentY - startY);
    const left = Math.min(currentX, startX);
    const top = Math.min(currentY, startY);

    selectionBox.style.left = `${left}px`;
    selectionBox.style.top = `${top}px`;
    selectionBox.style.width = `${width}px`;
    selectionBox.style.height = `${height}px`;
});

overlayContainer.addEventListener('mouseup', (e) => {
    if (!isDrawing) return;
    isDrawing = false;
    
    if (selectionBox) {
        selectionBox.style.display = 'none';
        
        const width = parseFloat(selectionBox.style.width);
        const height = parseFloat(selectionBox.style.height);
        
        // Ignore very small clicks
        if (width > 5 && height > 5) {
            const left = parseFloat(selectionBox.style.left);
            const top = parseFloat(selectionBox.style.top);
            
            addRedaction(left, top, width, height);
        }
    }
});

overlayContainer.addEventListener('mouseleave', () => {
    if (isDrawing) {
        isDrawing = false;
        if (selectionBox) selectionBox.style.display = 'none';
    }
});

function addRedaction(left, top, width, height) {
    // Convert DOM pixels to PDF points using currentViewport
    // currentViewport maps PDF coordinates to DOM pixels.
    // viewport.convertToPdfPoint(x, y) gives us [pdfX, pdfY]
    // Note: PDF point origin is bottom-left, but PDF.js convertToPdfPoint returns it relative to page?
    // Actually, it's easier to store the inverse of what we rendered:
    // If we drew at `left, top` in DOM, the scaled un-transformed coords are `left/scale, top/scale`.
    
    // Wait, PDF-lib origin is bottom-left.
    // Let's store coords relative to top-left in unscaled points.
    const pdfX = left / pdfScale;
    const pdfY = top / pdfScale;
    const pdfW = width / pdfScale;
    const pdfH = height / pdfScale;

    if (!redactions[currentPageNum]) {
        redactions[currentPageNum] = [];
    }
    
    const id = Date.now().toString();
    redactions[currentPageNum].push({ id, x: pdfX, y: pdfY, w: pdfW, h: pdfH });
    
    renderRedactionElements();
}

function removeRedaction(pageNum, id) {
    if (redactions[pageNum]) {
        redactions[pageNum] = redactions[pageNum].filter(r => r.id !== id);
        renderRedactionElements();
    }
}

function renderRedactionElements() {
    // Clear existing
    overlayContainer.innerHTML = '';
    if (selectionBox) overlayContainer.appendChild(selectionBox);
    
    const currentRedactions = redactions[currentPageNum] || [];
    
    currentRedactions.forEach(r => {
        const el = document.createElement('div');
        el.className = 'redact-rect';
        el.style.left = `${r.x * pdfScale}px`;
        el.style.top = `${r.y * pdfScale}px`;
        el.style.width = `${r.w * pdfScale}px`;
        el.style.height = `${r.h * pdfScale}px`;
        
        const delBtn = document.createElement('div');
        delBtn.className = 'delete-btn';
        delBtn.innerHTML = 'X';
        delBtn.onclick = (e) => {
            e.stopPropagation();
            removeRedaction(currentPageNum, r.id);
        };
        
        el.appendChild(delBtn);
        overlayContainer.appendChild(el);
    });
}

// Export logic
btnExport.addEventListener('click', async () => {
    showLoading('Aplicando tarjas...');
    
    try {
        // Load original PDF into pdf-lib
        const pdfDocMod = await PDFLib.PDFDocument.load(originalPdfBytes);
        
        // Ao invés de deletar as assinaturas, nós "achatamos" (flatten) o documento.
        // Isso transforma os selos flutuantes (anotações/formulários) em gráficos normais do fundo da página.
        // Assim, a assinatura original é preservada, mas nossas tarjas (que são desenhadas a seguir) 
        // ficarão corretamente por cima dela, caso você escolha tarjá-la.
        try {
            const form = pdfDocMod.getForm();
            if (form) {
                form.flatten();
            }
        } catch (e) {
            console.warn("Aviso ao achatar formulário:", e);
        }

        const pages = pdfDocMod.getPages();
        
        // Apply redactions
        for (const pageNumStr in redactions) {
            const pageNum = parseInt(pageNumStr, 10);
            const pageRedactions = redactions[pageNumStr];
            
            // pdf-lib pages are 0-indexed, pdf.js are 1-indexed
            const page = pages[pageNum - 1];
            const pageHeight = page.getHeight();
            // In pdf-lib, (0,0) is bottom-left. 
            // We have (x,y) from top-left.
            // newY = pageHeight - y - h
            
            for (const r of pageRedactions) {
                page.drawRectangle({
                    x: r.x,
                    y: pageHeight - r.y - r.h,
                    width: r.w,
                    height: r.h,
                    color: PDFLib.rgb(0, 0, 0),
                });
            }
        }
        
        // --- Registro no Painel de Auditoria (Power Automate) ---
        if (WEBHOOK_URL && WEBHOOK_URL.trim() !== '') {
            let totalRedactions = 0;
            for (const pageNumStr in redactions) {
                totalRedactions += redactions[pageNumStr].length;
            }
            
            try {
                fetch(WEBHOOK_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        usuarioResponsavel: currentUserName,
                        emailUsuario: currentUserEmail,
                        nomeDoArquivo: currentFileName,
                        numeroDePaginas: pages.length,
                        quantidadeDeTarjas: totalRedactions,
                        dataHora: new Date().toISOString()
                    })
                }).catch(e => console.warn("Erro ao registrar auditoria (rede):", e));
            } catch (e) {
                console.warn("Erro ao registrar auditoria (código):", e);
            }
        }
        // --------------------------------------------------------
        
        // Save and trigger download
        const modifiedPdfBytes = await pdfDocMod.save();
        
        const blob = new Blob([modifiedPdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'documento_tarjado.pdf';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showScreen(editorScreen);
        
    } catch (error) {
        console.error('Export error:', error);
        alert('Ocorreu um erro ao exportar o PDF: ' + error.message);
        showScreen(editorScreen);
    }
});
