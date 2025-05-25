// YouTube Playlist Extractor - Main JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const playlistForm = document.getElementById('playlist-form');
    const playlistUrlInput = document.getElementById('playlist-url');
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

    // Data storage
    let playlistData = {
        videos: [],
        screenshots: [],
        totalDuration: 0
    };

    // Event Listeners
    playlistForm.addEventListener('submit', handleFormSubmit);
    downloadExcelBtn.addEventListener('click', downloadExcel);
    downloadWordBtn.addEventListener('click', downloadWord);
    tryAgainBtn.addEventListener('click', resetForm);

    // Form submission handler
    function handleFormSubmit(e) {
        e.preventDefault();
        const playlistUrl = playlistUrlInput.value.trim();
        
        if (!isValidYouTubePlaylistUrl(playlistUrl)) {
            showError('Por favor, insira um link válido de playlist do YouTube.');
            return;
        }

        // Hide input section, show progress section
        document.querySelector('.input-section').classList.add('hidden');
        progressSection.classList.remove('hidden');
        
        // Extract playlist ID
        const playlistId = extractPlaylistId(playlistUrl);
        
        // Start extraction process
        extractPlaylistData(playlistId);
    }

    // Validate YouTube playlist URL
    function isValidYouTubePlaylistUrl(url) {
        const regex = /^(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/.*list=([a-zA-Z0-9_-]+)/;
        return regex.test(url);
    }

    // Extract playlist ID from URL
    function extractPlaylistId(url) {
        const regex = /list=([a-zA-Z0-9_-]+)/;
        const match = url.match(regex);
        return match ? match[1] : null;
    }

    // Main extraction function
    async function extractPlaylistData(playlistId) {
        try {
            // Update status
            progressStatus.textContent = 'Obtendo informações da playlist...';
            
            // Fetch playlist data using YouTube Data API
            const playlistItems = await fetchPlaylistItems(playlistId);
            
            if (!playlistItems || playlistItems.length === 0) {
                throw new Error('Não foi possível obter os vídeos da playlist.');
            }
            
            // Update UI with total videos
            videosTotal.textContent = playlistItems.length;
            
            // Process each video
            for (let i = 0; i < playlistItems.length; i++) {
                const videoItem = playlistItems[i];
                
                // Update progress
                const progress = Math.round(((i + 1) / playlistItems.length) * 100);
                progressBar.style.width = `${progress}%`;
                progressPercentage.textContent = `${progress}%`;
                videosProcessed.textContent = i + 1;
                currentVideoTitle.textContent = videoItem.title;
                progressStatus.textContent = `Processando vídeo ${i + 1} de ${playlistItems.length}...`;
                
                // Fetch detailed video data
                const videoData = await fetchVideoDetails(videoItem.videoId);
                
                // Take screenshot (simulated in this demo)
                const screenshot = await simulateTakeScreenshot(videoItem.videoId, videoItem.title);
                
                // Add to our data storage
                playlistData.videos.push(videoData);
                playlistData.screenshots.push(screenshot);
                playlistData.totalDuration += videoData.durationSeconds;
                
                // Small delay to show progress (can be removed in production)
                await new Promise(resolve => setTimeout(resolve, 500));
            }
            
            // All videos processed, show results
            showResults();
            
        } catch (error) {
            console.error('Extraction error:', error);
            showError(error.message || 'Ocorreu um erro durante a extração dos dados.');
        }
    }

    // Fetch playlist items from YouTube Data API
    async function fetchPlaylistItems(playlistId) {
        try {
            // In a real implementation, this would use the YouTube Data API
            // For this demo, we'll simulate the API response
            
            // Simulated delay
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            // Return simulated playlist items
            return simulatePlaylistItems(playlistId);
        } catch (error) {
            console.error('Error fetching playlist items:', error);
            throw new Error('Não foi possível acessar a playlist. Verifique o link e tente novamente.');
        }
    }

    // Fetch video details from YouTube Data API
    async function fetchVideoDetails(videoId) {
        try {
            // In a real implementation, this would use the YouTube Data API
            // For this demo, we'll simulate the API response
            
            // Simulated delay
            await new Promise(resolve => setTimeout(resolve, 800));
            
            // Return simulated video details
            return simulateVideoDetails(videoId);
        } catch (error) {
            console.error('Error fetching video details:', error);
            throw new Error(`Não foi possível obter detalhes do vídeo ${videoId}.`);
        }
    }

    // Simulate taking a screenshot (in a real app, this would use canvas or a similar approach)
    async function simulateTakeScreenshot(videoId, title) {
        // Simulated delay
        await new Promise(resolve => setTimeout(resolve, 600));
        
        // In a real implementation, this would capture a screenshot from the video
        // For this demo, we'll just return a placeholder
        return {
            videoId: videoId,
            title: title,
            dataUrl: `https://img.youtube.com/vi/${videoId}/maxresdefault.jpg`,
            timestamp: new Date().toISOString()
        };
    }

    // Show results section
    function showResults() {
        progressSection.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        
        // Update stats
        totalVideosSpan.textContent = playlistData.videos.length;
        totalScreenshotsSpan.textContent = playlistData.screenshots.length;
        totalTimeSpan.textContent = formatDuration(playlistData.totalDuration);
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
        playlistData = {
            videos: [],
            screenshots: [],
            totalDuration: 0
        };
    }

    // Download Excel file
    function downloadExcel() {
        // In a real implementation, this would generate an Excel file
        // For this demo, we'll simulate the download
        
        progressStatus.textContent = 'Gerando planilha Excel...';
        
        // Simulate processing time
        setTimeout(() => {
            // Create a sample Excel file (in a real app, use a library like SheetJS)
            const blob = new Blob(['Sample Excel content'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            
            // Create download link
            const a = document.createElement('a');
            a.href = url;
            a.download = 'playlist_data.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            // Clean up
            URL.revokeObjectURL(url);
        }, 1000);
    }

    // Download Word document
    function downloadWord() {
        // In a real implementation, this would generate a Word document
        // For this demo, we'll simulate the download
        
        progressStatus.textContent = 'Gerando documento Word...';
        
        // Simulate processing time
        setTimeout(() => {
            // Create a sample Word file (in a real app, use a library like docx)
            const blob = new Blob(['Sample Word content'], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            const url = URL.createObjectURL(blob);
            
            // Create download link
            const a = document.createElement('a');
            a.href = url;
            a.download = 'playlist_screenshots.docx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            // Clean up
            URL.revokeObjectURL(url);
        }, 1000);
    }

    // Helper function to format duration in seconds to HH:MM:SS
    function formatDuration(seconds) {
        const hours = Math.floor(seconds / 3600);
        const minutes = Math.floor((seconds % 3600) / 60);
        const secs = seconds % 60;
        
        return `${hours}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    }

    // Simulation functions (these would be replaced with actual API calls in a real implementation)
    function simulatePlaylistItems(playlistId) {
        // Generate 5-15 simulated videos
        const count = Math.floor(Math.random() * 11) + 5;
        const items = [];
        
        for (let i = 0; i < count; i++) {
            const videoId = generateRandomVideoId();
            items.push({
                videoId: videoId,
                title: `Vídeo ${i + 1} da Playlist ${playlistId.substring(0, 6)}`,
                position: i,
                thumbnailUrl: `https://img.youtube.com/vi/${videoId}/default.jpg`
            });
        }
        
        return items;
    }

    function simulateVideoDetails(videoId) {
        // Generate random video details
        const views = Math.floor(Math.random() * 1000000) + 1000;
        const likes = Math.floor(Math.random() * 50000) + 100;
        const durationSeconds = Math.floor(Math.random() * 900) + 120; // 2-17 minutes
        const publishedDate = new Date(Date.now() - Math.floor(Math.random() * 31536000000)); // Within last year
        
        return {
            videoId: videoId,
            title: `Vídeo ${videoId.substring(0, 6)}`,
            views: views,
            likes: likes,
            durationSeconds: durationSeconds,
            duration: formatDuration(durationSeconds),
            publishedDate: publishedDate.toISOString(),
            channelTitle: 'Canal de Exemplo',
            description: 'Esta é uma descrição de exemplo para o vídeo simulado.'
        };
    }

    function generateRandomVideoId() {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
        let result = '';
        for (let i = 0; i < 11; i++) {
            result += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        return result;
    }
});
