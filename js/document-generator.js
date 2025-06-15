// File: js/app.js

// Integration of UI with YouTube Extractor and Document Generator including Screenshot Token
document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const playlistForm = document.getElementById('playlist-form');
    const playlistUrlInput = document.getElementById('playlist-url');
    const apiKeyInput = document.getElementById('api-key');
    const screenshotTokenInput = document.getElementById('screenshot-token');
    const extractBtn = document.getElementById('extract-btn');
    const progressSection = document.getElementById('progress-section');
    const resultsSection = document.getElementById('results-section');
    const errorSection = document.getElementById('error-section');
    const progressBar = document.getElementById('progress-bar');
    const progressPercentage = document.getElementById('progress-percentage');
    const progressStatus = document.getElementById('progress-status');
    const videosProcessed = document.getElementById('videos-processed');
    const videosTotal = document.getElementById('total-videos');
    const currentVideoTitle = document.getElementById('current-video-title');
    const totalVideosSpan = document.getElementById('total-videos');
    const totalScreenshotsSpan = document.getElementById('total-screenshots');
    const totalTimeSpan = document.getElementById('total-time');
    const downloadExcelBtn = document.getElementById('download-excel');
    const downloadWordBtn = document.getElementById('download-word');
    const tryAgainBtn = document.getElementById('try-again-btn');
    const errorMessage = document.getElementById('error-message');
    const apiInstructionsLink = document.getElementById('api-instructions-link');
    const apiModal = document.getElementById('api-modal');
    const closeModal = document.querySelector('.close-modal');

    // Load stored API keys
    const storedYtKey = localStorage.getItem('youtube_api_key');
    if (storedYtKey) apiKeyInput.value = storedYtKey;
    const storedScreenshotToken = localStorage.getItem('screenshot_api_token');
    if (storedScreenshotToken) screenshotTokenInput.value = storedScreenshotToken;

    // Save token fields on change
    apiKeyInput.addEventListener('change', () => {
        const val = apiKeyInput.value.trim();
        if (val) localStorage.setItem('youtube_api_key', val);
        else localStorage.removeItem('youtube_api_key');
    });
    screenshotTokenInput.addEventListener('change', () => {
        const val = screenshotTokenInput.value.trim();
        if (val) localStorage.setItem('screenshot_api_token', val);
        else localStorage.removeItem('screenshot_api_token');
    });

    // Initialize extractor and document generator with token
    const extractor = new YouTubeExtractor();
    const docGenerator = new DocumentGenerator(screenshotTokenInput.value.trim());

    // Data storage for extracted information
    let extractedData = { videos: [], screenshots: [], totalDuration: 0 };

    // Event Listeners
    playlistForm.addEventListener('submit', handleFormSubmit);
    downloadExcelBtn.addEventListener('click', handleExcelDownload);
    downloadWordBtn.addEventListener('click', handleWordDownload);
    tryAgainBtn.addEventListener('click', resetForm);

    // Modal handling
    if (apiInstructionsLink && apiModal) {
        apiInstructionsLink.addEventListener('click', function(e) {
            e.preventDefault(); apiModal.style.display = 'flex';
        });
    }
    if (closeModal && apiModal) {
        closeModal.addEventListener('click', () => apiModal.style.display = 'none');
    }
    window.addEventListener('click', event => {
        if (event.target === apiModal) apiModal.style.display = 'none';
    });

    async function handleFormSubmit(e) {
        e.preventDefault();
        const playlistUrl = playlistUrlInput.value.trim();
        const apiKey = apiKeyInput.value.trim();
        const screenshotToken = screenshotTokenInput.value.trim();

        if (!isValidYouTubePlaylistUrl(playlistUrl)) {
            showError('Por favor, insira um link válido de playlist do YouTube.');
            return;
        }
        // Save keys (change event already handled)
        extractor.setApiKey(apiKey);
        docGenerator.setScreenshotToken(screenshotToken);

        // UI
        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.remove('hidden');

        const playlistId = extractPlaylistId(playlistUrl);
        try {
            extractedData = await extractor.processPlaylist(playlistId, updateProgress);
            showResults();
        } catch (error) {
            console.error('Extraction error:', error);
            showError(error.message || 'Erro durante a extração.');
        }
    }

    function updateProgress(progress) {
        progressBar.style.width = `${progress.progress}%`;
        progressPercentage.textContent = `${progress.progress}%`;
        progressStatus.textContent = progress.status;
        if (progress.total > 0) {
            videosTotal.textContent = progress.total;
            videosProcessed.textContent = progress.processed;
        }
        if (progress.currentVideo) currentVideoTitle.textContent = progress.currentVideo;
    }

    function isValidYouTubePlaylistUrl(url) {
        return /^(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/.*list=[\w-]+/.test(url);
    }
    function extractPlaylistId(url) {
        const match = url.match(/list=([\w-]+)/);
        return match ? match[1] : null;
    }
    function showResults() {
        progressSection.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        totalVideosSpan.textContent = extractedData.videos.length;
        totalScreenshotsSpan.textContent = extractedData.screenshots.length;
        totalTimeSpan.textContent = formatDuration(extractedData.totalDuration);
    }
    function showError(msg) {
        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        errorSection.classList.remove('hidden');
        errorMessage.textContent = msg;
    }
    function resetForm() {
        document.querySelector('.input-section').classList.remove('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        progressBar.style.width = '0%'; progressPercentage.textContent = '0%';
        progressStatus.textContent = 'Iniciando...'; videosProcessed.textContent='0'; videosTotal.textContent='0'; currentVideoTitle.textContent='Aguardando...';
        extractedData = { videos: [], screenshots: [], totalDuration: 0 };
    }
    async function handleExcelDownload() {
        if (!extractedData.videos.length) return showError('Nenhum dado para download.');
        try {
            progressStatus.textContent = 'Gerando planilha Excel...';
            const excelData = docGenerator.generateExcel(extractedData.videos);
            downloadFile(excelData.blob, excelData.filename);
            progressStatus.textContent = 'Excel baixado!';
        } catch (e) { console.error(e); progressStatus.textContent='Erro no Excel.'; }
    }
    async function handleWordDownload() {
        if (!extractedData.videos.length) return showError('Nenhum dado para download.');
        try {
            progressStatus.textContent = 'Gerando documento...';
            const wordData = await docGenerator.generateWordDocument(extractedData.videos);
            downloadFile(wordData.blob, wordData.filename);
            progressStatus.textContent='Documento baixado!';
        } catch (e) { console.error(e); progressStatus.textContent='Erro no documento.'; }
    }
    function downloadFile(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
    function formatDuration(s) {
        const h = Math.floor(s/3600), m=Math.floor((s%3600)/60), ss=s%60;
        return h?`${h}:${String(m).padStart(2,'0')}:${String(ss).padStart(2,'0')}`:`${m}:${String(ss).padStart(2,'0')}`;
    }
});


// File: js/document-generator.js

// Document Generator with DOCX Support and direct Screenshot API calls
class DocumentGenerator {
  constructor(screenshotToken) {
    this.xlsx = null;
    this.docx = null;
    this.fileSaver = null;
    this.screenshotToken = screenshotToken;
    this.loadLibraries();
  }

  setScreenshotToken(token) {
    this.screenshotToken = token;
  }

  async loadLibraries() {
    if (!window.XLSX) await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    this.xlsx = window.XLSX;
    if (!window.docx) await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js');
    this.docx = window.docx;
    if (!window.saveAs) await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
    this.fileSaver = window.saveAs;
  }

  loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s);
    });
  }

  generateExcel(videos) {
    if (!videos.length) throw new Error('Sem dados para Excel');
    const data = videos.map(v => ({ 'Nome do Episodio': v.title, 'Duração': v.duration, 'Views': v.views, 'Likes': v.likes, 'Link':`https://www.youtube.com/watch?v=${v.videoId}`, 'Data de Publicacao': new Date(v.publishedDate).toLocaleDateString() }));
    const ws = this.xlsx.utils.json_to_sheet(data);
    const wb = this.xlsx.utils.book_new(); this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');
    const buf = this.xlsx.write(wb, {bookType:'xlsx',type:'array'});
    return { blob: new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), filename:'playlist_data.xlsx' };
  }

  async generateWordDocument(videos) {
    if (!videos.length) throw new Error('Sem dados para doc');
    if (!this.docx) return this.generateHtmlDocument(videos);
    const { Document, Paragraph, Table, TableRow, TableCell, BorderStyle, AlignmentType, HeadingLevel, PageBreak, ImageRun, Packer } = this.docx;
    const doc = new Document({ creator:'YouTube Extractor', title:'Comprovação de Dados', description:'Playlist com screenshots' });
    doc.addSection({properties:{}, children:[ new Paragraph({text:'Comprovação de Dados - Playlist', heading:HeadingLevel.HEADING_1, alignment:AlignmentType.CENTER}) ]});
    for (let i=0; i<videos.length; i++) {
      const v = videos[i]; const url=`https://www.youtube.com/watch?v=${v.videoId}`;
      const children=[ new Paragraph({text:`Vídeo ${i+1}: ${v.title}`, heading:HeadingLevel.HEADING_2}) ];
      children.push(new Table({rows:[ this._row('Duração',v.duration), this._row('Views',String(v.views)), this._row('Likes',String(v.likes)), this._row('Link',url) ], width:{size:100,type:'pct'}, borders:{top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}}}));
      let imgData=null;
      if (this.screenshotToken) {
        try {
          imgData = await this._fetchScreenshot(url);
        } catch(e){console.warn('Screenshot failed',e);}  
      }
      if (imgData) children.push(new Paragraph({children:[new ImageRun({data:imgData.split(',')[1],transformation:{width:600,height:338}})]}));
      if (i<videos.length-1) children.push(new Paragraph({children:[new PageBreak()]}));
      doc.addSection({properties:{},children});
    }
    const buffer=await Packer.toBuffer(doc);
    return { blob:new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}), filename:'comprovacao_videos.docx' };
  }

  _row(label,val){ const { TableRow, TableCell, Paragraph } = this.docx; return new TableRow({children:[ new TableCell({width:{size:30,type:'pct'},children:[new Paragraph({text:label+':' ,bold:true})]}), new TableCell({width:{size:70,type:'pct'},children:[new Paragraph({text:val})]}) ]}); }

  async _fetchScreenshot(url) {
    const api=`https://api.screenshotapi.net/screenshot?token=${this.screenshotToken}&url=${encodeURIComponent(url)}&full_page=true`;
    const res=await fetch(api); const json=await res.json(); return `data:image/png;base64,${json.screenshot}`;
  }

  generateHtmlDocument(videos){ let html='<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação</title></head><body>'; videos.forEach((v,i)=>{ html+=`<h2>Vídeo ${i+1}: ${v.title}</h2><ul><li>Duração: ${v.duration}</li><li>Views: ${v.views}</li><li>Likes: ${v.likes}</li><li>Link: <a href="https://www.youtube.com/watch?v=${v.videoId}">YouTube</a></li></ul><img src="https://img.youtube.com/vi/${v.videoId}/maxresdefault.jpg" style="max-width:600px;"><hr>`}); html+='</body></html>'; return { blob:new Blob([html],{type:'text/html;charset=utf-8;'}), filename:'comprovacao_videos.html' }; }
}
