// Integration of UI with YouTube Extractor and Document Generator
document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const playlistForm = document.getElementById('playlist-form');
    const playlistUrlInput = document.getElementById('playlist-url');
    const apiKeyInput = document.getElementById('api-key');
    const extractBtn = document.getElementById('extract-btn');
    const progressSection = document.getElementById('progress-section');
    const resultsSection = document.getElementById('results-section');
    const errorSection = document.getElementById('error-section');
    const progressBar = document.getElementById('progress-bar');
    const progressPercentage = document.getElementById('progress-percentage');
    const progressStatus = document.getElementById('progress-status');
    const videosProcessed = document.getElementById('videos-processed');
    const videosTotal = document.getElementById('videos-total');
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

    // Initialize extractor and document generator
    const extractor = new YouTubeExtractor();
    const docGenerator = new DocumentGenerator();
    
    // Data storage for extracted information
    let extractedData = {
        videos: [],
        screenshots: [],
        totalDuration: 0
    };

    // Check for saved API key in localStorage
    if (localStorage.getItem('youtube_api_key')) {
        apiKeyInput.value = localStorage.getItem('youtube_api_key');
    }

    // Event Listeners
    playlistForm.addEventListener('submit', handleFormSubmit);
    downloadExcelBtn.addEventListener('click', handleExcelDownload);
    downloadWordBtn.addEventListener('click', handleWordDownload);
    tryAgainBtn.addEventListener('click', resetForm);
    
    // Modal handling - FIXED VERSION
    if (apiInstructionsLink) {
        apiInstructionsLink.addEventListener('click', function(e) {
            e.preventDefault();
            if (apiModal) {
                apiModal.style.display = 'flex';
            }
        });
    }
    
    if (closeModal && apiModal) {
        closeModal.addEventListener('click', function() {
            apiModal.style.display = 'none';
        });
    }
    
    // Close modal when clicking outside of it
    window.addEventListener('click', function(event) {
        if (event.target === apiModal) {
            apiModal.style.display = 'none';
        }
    });

    // Form submission handler
    async function handleFormSubmit(e) {
        e.preventDefault();
        const playlistUrl = playlistUrlInput.value.trim();
        const apiKey = apiKeyInput.value.trim();
        
        if (!isValidYouTubePlaylistUrl(playlistUrl)) {
            showError('Por favor, insira um link válido de playlist do YouTube.');
            return;
        }

        // Save API key to localStorage if provided
        if (apiKey) {
            localStorage.setItem('youtube_api_key', apiKey);
            extractor.setApiKey(apiKey);
        } else {
            localStorage.removeItem('youtube_api_key');
        }

        // Hide input section, show progress section
        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.remove('hidden');
        
        // Extract playlist ID
        const playlistId = extractPlaylistId(playlistUrl);
        
        try {
            // Start extraction process
            extractedData = await extractor.processPlaylist(playlistId, updateProgress);
            
            // Show results
            showResults();
        } catch (error) {
            console.error('Extraction error:', error);
            showError(error.message || 'Ocorreu um erro durante a extração dos dados.');
        }
    }

    // Progress update callback
    function updateProgress(progress) {
        progressBar.style.width = `${progress.progress}%`;
        progressPercentage.textContent = `${progress.progress}%`;
        progressStatus.textContent = progress.status;
        
        if (progress.total > 0) {
            videosTotal.textContent = progress.total;
            videosProcessed.textContent = progress.processed;
        }
        
        if (progress.currentVideo) {
            currentVideoTitle.textContent = progress.currentVideo;
        }
    }

    // Validate YouTube playlist URL - IMPROVED VERSION
    function isValidYouTubePlaylistUrl(url) {
        // Expressão regular melhorada para aceitar parâmetros adicionais após o ID da playlist
        const regex = /^(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/.*list=([a-zA-Z0-9_-]+)/;
        return regex.test(url);
    }

    // Extract playlist ID from URL - IMPROVED VERSION
    function extractPlaylistId(url) {
        // Expressão regular para extrair apenas o ID da playlist, ignorando parâmetros adicionais
        const regex = /list=([a-zA-Z0-9_-]+)/;
        const match = url.match(regex);
        return match ? match[1] : null;
    }

    // Show results section
    function showResults() {
        progressSection.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        
        // Update stats
        totalVideosSpan.textContent = extractedData.videos.length;
        totalScreenshotsSpan.textContent = extractedData.screenshots.length;
        totalTimeSpan.textContent = formatDuration(extractedData.totalDuration);
    }

    // Show error section
    function showError(message) {
        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        errorSection.classList.remove('hidden');
        
        errorMessage.textContent = message;
    }

    // Reset form to initial state
    function resetForm() {
        document.querySelector('.input-section').classList.remove('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        
        // Reset progress
        progressBar.style.width = '0%';
        progressPercentage.textContent = '0%';
        progressStatus.textContent = 'Iniciando...';
        videosProcessed.textContent = '0';
        videosTotal.textContent = '0';
        currentVideoTitle.textContent = 'Aguardando...';
        
        // Clear data
        extractedData = {
            videos: [],
            screenshots: [],
            totalDuration: 0
        };
    }

    // Handle Excel download
    async function handleExcelDownload() {
        if (extractedData.videos.length === 0) {
            showError('Nenhum dado disponível para download.');
            return;
        }
        
        try {
            progressStatus.textContent = 'Gerando planilha Excel...';
            
            // Generate Excel file
            const excelData = docGenerator.generateExcel(extractedData.videos);
            
            // Create download link
            downloadFile(excelData.blob, excelData.filename);
            
            progressStatus.textContent = 'Planilha Excel baixada com sucesso!';
        } catch (error) {
            console.error('Error downloading Excel:', error);
            progressStatus.textContent = 'Erro ao gerar planilha Excel.';
        }
    }

    // Handle Word document download
    async function handleWordDownload() {
        if (extractedData.videos.length === 0 || extractedData.screenshots.length === 0) {
            showError('Nenhum dado disponível para download.');
            return;
        }
        
        try {
            progressStatus.textContent = 'Gerando documento com capturas...';
            
            // Generate Word document
            const wordData = docGenerator.generateWordDocument(extractedData.videos, extractedData.screenshots);
            
            // Create download link
            downloadFile(wordData.blob, wordData.filename);
            
            progressStatus.textContent = 'Documento baixado com sucesso!';
        } catch (error) {
            console.error('Error downloading Word document:', error);
            progressStatus.textContent = 'Erro ao gerar documento.';
        }
    }

    // Helper function to download a file
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

    // Helper function to format duration in seconds to HH:MM:SS
    function formatDuration(seconds) {
        const hours = Math.floor(seconds / 3600);
        const minutes = Math.floor((seconds % 3600) / 60);
        const secs = seconds % 60;
        
        if (hours > 0) {
            return `${hours}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
        } else {
            return `${minutes}:${secs.toString().padStart(2, '0')}`;
        }
    }
});
