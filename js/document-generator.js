// Full document-generator.js with all methods
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
      if (!window.XLSX) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
      }
      this.xlsx = window.XLSX;

      try {
        await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.4.0/build/index.umd.js');
        if (window.docx && window.docx.Packer && typeof window.docx.Packer.toBuffer === 'function') {
          this.docx = window.docx;
          this.docxLoaded = true;
        } else {
          console.warn('docx library failed to load correctly.');
        }
      } catch (docxErr) {
        console.warn('Could not load docx library from CDN:', docxErr);
      }

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
      if (this.xlsx) {
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

        return {
          blob,
          filename: 'playlist_data.xlsx'
        };
      } else {
        return this.generateCSV(videos);
      }
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

    return {
      blob,
      filename: 'playlist_data.csv'
    };
  }

  async generateWordDocument(videos, screenshots) {
    if (!videos || videos.length === 0 || !screenshots || screenshots.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }

    try {
      if (!this.docxLoaded) {
        console.warn('DOCX library not loaded or invalid, falling back to HTML');
        return this.generateHtmlDocument(videos, screenshots);
      }

      const { Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, Packer, HeadingLevel } = this.docx;

      const doc = new Document({
  creator: "Lavvi",
  title: "Comprovação de Dados - Playlist YouTube",
  description: "Documento gerado automaticamente com os dados da playlist.",
  sections: [] });

      

      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];
        const screenshot = screenshots[i];
        const imageData = await this.getImageDataFromUrl(screenshot.imageUrl);

        doc.sections.push({ properties: {}, children: [
            new Paragraph({
              text: `Vídeo ${i + 1}: ${video.title}`,
              heading: HeadingLevel.HEADING_2
            }),
            new Table({
              rows: [
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Nome do Episódio:')] }),
                  new TableCell({ children: [new Paragraph(video.title)] })
                ] }),
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Duração:')] }),
                  new TableCell({ children: [new Paragraph(video.duration)] })
                ] }),
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Views:')] }),
                  new TableCell({ children: [new Paragraph(video.views.toLocaleString())] })
                ] }),
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Likes:')] }),
                  new TableCell({ children: [new Paragraph(video.likes.toLocaleString())] })
                ] }),
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Link:')] }),
                  new TableCell({ children: [new Paragraph(`https://www.youtube.com/watch?v=${video.videoId}`)] })
                ] }),
                new TableRow({ children: [
                  new TableCell({ children: [new Paragraph('Data de Publicação:')] }),
                  new TableCell({ children: [new Paragraph(new Date(video.publishedDate).toLocaleDateString())] })
                ] })
              ]
            }),
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageData,
                  transformation: { width: 600, height: 350 }
                })
              ]
            })
          ]
        });
      }

      const buffer = await Packer.toBuffer(doc);
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

      return {
        blob,
        filename: 'comprovacao_videos.docx'
      };
    } catch (error) {
      console.error('Error generating Word document:', error);
      return this.generateHtmlDocument(videos, screenshots);
    }
  }

  async getImageDataFromUrl(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL('image/png');
        const base64 = dataUrl.split(',')[1];
        resolve(base64);
      };
      img.onerror = () => reject(new Error('Failed to load image'));
      img.src = url;
    });
  }

  generateHtmlDocument(videos, screenshots) {
    let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação</title></head><body><h1>Comprovação de Dados</h1>`;
    videos.forEach((video, i) => {
      html += `<h2>${video.title}</h2><ul>` +
              `<li>Duração: ${video.duration}</li>` +
              `<li>Views: ${video.views.toLocaleString()}</li>` +
              `<li>Likes: ${video.likes.toLocaleString()}</li>` +
              `<li>Publicado em: ${new Date(video.publishedDate).toLocaleDateString()}</li>` +
              `<li><a href="https://www.youtube.com/watch?v=${video.videoId}">Link</a></li>` +
              `</ul><img src="${screenshots[i].imageUrl}" style="max-width:100%;"><hr/>`;
    });
    html += '</body></html>';
    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    return { blob, filename: 'comprovacao_videos.html' };
  }
}
