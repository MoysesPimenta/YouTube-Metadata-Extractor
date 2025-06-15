// File: js/app.js

// Integration of UI with YouTube Extractor and Document Generator including Screenshot Token
document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const playlistForm = document.getElementById('playlist-form');
    const playlistUrlInput = document.getElementById('playlist-url');
    const apiKeyInput = document.getElementById('api-key');
    const screenshotTokenInput = document.getElementById('screenshot-token');
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
    let extractedData = { videos: [], totalDuration: 0 };

    // Event Listeners
    playlistForm.addEventListener('submit', handleFormSubmit);
    downloadExcelBtn.addEventListener('click', handleExcelDownload);
    downloadWordBtn.addEventListener('click', handleWordDownload);
    tryAgainBtn.addEventListener('click', resetForm);

    async function handleFormSubmit(e) {
        e.preventDefault();
        const playlistUrl = playlistUrlInput.value.trim();
        const apiKey = apiKeyInput.value.trim();
        const screenshotToken = screenshotTokenInput.value.trim();

        if (!/^(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/.*list=[\w-]+/.test(playlistUrl)) {
            showError('Por favor, insira um link válido de playlist do YouTube.');
            return;
        }

        extractor.setApiKey(apiKey);
        docGenerator.setScreenshotToken(screenshotToken);

        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.remove('hidden');

        const playlistId = (playlistUrl.match(/list=([\w-]+)/) || [])[1];
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
    }

    function showResults() {
        progressSection.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        totalVideosSpan.textContent = extractedData.videos.length;
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
        progressBar.style.width = '0%';
        progressPercentage.textContent = '0%';
        progressStatus.textContent = 'Iniciando...';
        extractedData = { videos: [], totalDuration: 0 };
    }

    async function handleExcelDownload() {
        if (!extractedData.videos.length) return showError('Nenhum dado para download.');
        try {
            const excelData = docGenerator.generateExcel(extractedData.videos);
            downloadFile(excelData.blob, excelData.filename);
        } catch (e) {
            console.error(e);
            showError('Erro ao gerar Excel.');
        }
    }

    async function handleWordDownload() {
        if (!extractedData.videos.length) return showError('Nenhum dado para download.');
        try {
            const wordData = await docGenerator.generateWordDocument(extractedData.videos);
            downloadFile(wordData.blob, wordData.filename);
        } catch (e) {
            console.error(e);
            showError('Erro ao gerar documento.');
        }
    }

    function downloadFile(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function formatDuration(sec) {
        const h = Math.floor(sec / 3600);
        const m = Math.floor((sec % 3600) / 60);
        const s = sec % 60;
        return h ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}`;
    }
});
