// Full, final document-generator.js
class DocumentGenerator {
  constructor() {
    this.xlsx = null;
    this.docx = null;
    this.fileSaver = null;
    this.docxLoaded = false;
    this.loadLibraries();
  }

  async loadLibraries() {
    try {
      // Load SheetJS for Excel
      if (!window.XLSX) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
      }
      this.xlsx = window.XLSX;

      // Load docx library for Word
      try {
        await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.4.0/build/index.umd.js');
        if (window.docx && window.docx.Packer && typeof window.docx.Packer.toBlob === 'function') {
          this.docx = window.docx;
          this.docxLoaded = true;
        } else {
          console.warn('docx library loaded but missing browser support.');
        }
      } catch (err) {
        console.warn('Failed to load docx library:', err);
      }

      // Load FileSaver
      if (!window.saveAs) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
      }
      this.fileSaver = window.saveAs;

      console.log('Libraries loaded successfully');
    } catch (error) {
      console.error('Error loading libraries:', error);
    }
  }

  loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.onload = resolve;
      script.onerror = reject;
      document.head.appendChild(script);
    });
  }

  generateExcel(videos) {
    if (!videos || videos.length === 0) {
      throw new Error('Nenhum dado disponível para gerar a planilha.');
    }
    try {
      const data = videos.map(video => ({
        'Nome do Episodio': video.title,
        'Duração': video.duration,
        'Views': video.views,
        'Likes': video.likes,
        'Link': `https://www.youtube.com/watch?v=${video.videoId}`,
        'Data de Publicacao': new Date(video.publishedDate).toLocaleDateString()
      }));
      const ws = this.xlsx.utils.json_to_sheet(data);
      const wb = this.xlsx.utils.book_new();
      this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');
      const excelBuffer = this.xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return { blob, filename: 'playlist_data.xlsx' };
    } catch (error) {
      console.error('Error generating Excel:', error);
      return this.generateCSV(videos);
    }
  }

  generateCSV(videos) {
    let csv = 'Nome do Episodio,Duração,Views,Likes,Link,Data de Publicacao\n';
    videos.forEach(video => {
      const row = [
        `"${video.title.replace(/"/g, '""')}"`,
        video.duration,
        video.views,
        video.likes,
        `"https://www.youtube.com/watch?v=${video.videoId}"`,
        new Date(video.publishedDate).toLocaleDateString()
      ];
      csv += row.join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    return { blob, filename: 'playlist_data.csv' };
  }

  async generateWordDocument(videos, screenshots) {
    if (!videos || videos.length === 0 || !screenshots || screenshots.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }
    // Fallback: generate HTML and save as .doc for Word compatibility
    const htmlDoc = this.generateHtmlDocument(videos, screenshots);
    // Convert HTML blob to Word-compatible blob
    const arrayBuffer = await htmlDoc.blob.arrayBuffer();
    const wordBlob = new Blob([arrayBuffer], { type: 'application/msword' });
    return { blob: wordBlob, filename: 'comprovacao_videos.doc' };
  }

  async getImageDataFromUrl(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      img.onload = () => {
        const canvas = document.createElement('canvas'); canvas.width = img.width; canvas.height = img.height;
        const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL('image/png');
        resolve(dataUrl.split(',')[1]);
      };
      img.onerror = () => reject();
      img.src = url;
    });
  }

  generateHtmlDocument(videos, screenshots) {
    let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação</title></head><body><h1>Comprovação de Dados</h1>';
    videos.forEach((video,i) =>{
      html += `<h2>${video.title}</h2><ul>`+
        `<li>Duração: ${video.duration}</li>`+
        `<li>Views: ${video.views.toLocaleString()}</li>`+
        `<li>Likes: ${video.likes.toLocaleString()}</li>`+
        `<li>Publicado em: ${new Date(video.publishedDate).toLocaleDateString()}</li>`+
        `<li><a href="https://www.youtube.com/watch?v=${video.videoId}">Link</a></li></ul>`+
        `<img src="${screenshots[i].imageUrl}" style="max-width:100%;"><hr/>`;
    });
    html += '</body></html>';
    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    return { blob, filename: 'comprovacao_videos.html' };
  }
}

