// Document Generator with DOCX Support
class DocumentGenerator {
  constructor() {
    this.xlsx = null;
    this.docx = null;
    this.fileSaver = null;
    this.loadLibraries();
  }

  // Load required libraries
  async loadLibraries() {
    try {
      // Load SheetJS (xlsx) for Excel generation
      if (!window.XLSX) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
      }
      this.xlsx = window.XLSX;

      // Load docx.js for Word document generation
      if (!window.docx) {
        // Use the official Docx CDN (UMD bundle)
        await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js');
      }
      this.docx = window.docx;

      // Load FileSaver for saving files
      if (!window.saveAs) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
      }
      this.fileSaver = window.saveAs;

      console.log('Libraries loaded successfully');
    } catch (error) {
      console.error('Error loading libraries:', error);
    }
  }

  // Helper to load external scripts
  loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.onload = resolve;
      script.onerror = reject;
      document.head.appendChild(script);
    });
  }

  // Generate Excel file from video data
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

  // Fallback: Generate CSV file
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

  // Generate Word document (DOCX format)
  async generateWordDocument(videos, screenshots) {
    if (!videos || videos.length === 0 || !screenshots || screenshots.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }

    try {
      if (!this.docx) {
        console.warn('DOCX library not loaded, falling back to HTML');
        return this.generateHtmlDocument(videos, screenshots);
      }
      // Destructure all needed components, including Packer
      const {
        Document,
        Paragraph,
        TextRun,
        ImageRun,
        Table,
        TableRow,
        TableCell,
        BorderStyle,
        AlignmentType,
        HeadingLevel,
        PageBreak,
        Packer
      } = this.docx;

      // Create the Word document
      const doc = new Document({
        creator: 'YouTube Playlist Extractor',
        title: 'Comprovação de Dados - Playlist YouTube',
        description: 'Documento de comprovação de dados extraídos de uma playlist do YouTube',
        styles: {
          paragraphStyles: [
            {
              id: 'Heading1',
              name: 'Heading 1',
              basedOn: 'Normal',
              next: 'Normal',
              quickFormat: true,
              run: {
                size: 48,
                bold: true
              },
              paragraph: {
                spacing: { after: 240 }
              }
            },
            {
              id: 'MetadataLabel',
              name: 'Metadata Label',
              basedOn: 'Normal',
              quickFormat: true,
              run: {
                color: '555555',
                italics: true
              }
            },
            {
              id: 'MetadataValue',
              name: 'Metadata Value',
              basedOn: 'Normal',
              quickFormat: true,
              run: {
                color: '000000'
              }
            },
            {
              id: 'HighlightedValue',
              name: 'Highlighted Value',
              basedOn: 'Normal',
              quickFormat: true,
              run: {
                color: '0000FF',
                underline: {}
              }
            }
          ]
        }
      });

      // Title section
      doc.addSection({
        properties: {},
        children: [
          new Paragraph({
            text: 'Comprovação de Dados - Playlist YouTube',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER
          })
        ]
      });

      // Iterate videos and screenshots
      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];
        const screenshot = screenshots[i];
        let imageData = null;
        try {
          imageData = await this.getImageDataFromUrl(screenshot.imageUrl);
        } catch (e) {
          console.warn('Failed to fetch image for DOCX:', e);
        }

        // Build section for each video
        const sectionChildren = [];

        // Video header
        sectionChildren.push(new Paragraph({
          text: `Vídeo ${i + 1}: ${video.title}`,
          heading: HeadingLevel.HEADING_2
        }));

        // Metadata table
        sectionChildren.push(new Table({
          rows: [
            // Episode Name
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Nome do Episódio:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [ new Paragraph({ text: video.title, style: 'MetadataValue' }) ]
                })
              ]
            }),
            // Duration
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Duração:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [ new Paragraph({ text: video.duration, style: 'MetadataValue' }) ]
                })
              ]
            }),
            // Views
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Views:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [ new Paragraph({ text: video.views.toString(), style: 'MetadataValue' }) ]
                })
              ]
            }),
            // Likes
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Likes:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [ new Paragraph({ text: video.likes.toString(), style: 'MetadataValue' }) ]
                })
              ]
            }),
            // Link (highlighted)
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Link:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({ text: `https://www.youtube.com/watch?v=${video.videoId}`, style: 'HighlightedValue' })
                      ]
                    })
                  ]
                })
              ]
            }),
            // Publication Date
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 30, type: 'pct' },
                  children: [ new Paragraph({ text: 'Data Publicação:', style: 'MetadataLabel' }) ]
                }),
                new TableCell({
                  width: { size: 70, type: 'pct' },
                  children: [ new Paragraph({ text: new Date(video.publishedDate).toLocaleDateString(), style: 'MetadataValue' }) ]
                })
              ]
            })
          ],
          width: { size: 100, type: 'pct' },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideHorizontal: { style: BorderStyle.NONE },
            insideVertical: { style: BorderStyle.NONE }
          }
        }));

        // Insert screenshot image if available
        if (imageData) {
          sectionChildren.push(new Paragraph({
            children: [
              new ImageRun({
                data: imageData.split(',')[1], // base64 data after comma
                transformation: { width: 600, height: 338 }
              })
            ]
          }));
        }

        // Page break after each video (except last)
        if (i < videos.length - 1) {
          sectionChildren.push(new Paragraph({ children: [ new PageBreak() ] }));
        }

        // Add section to document
        doc.addSection({ properties: {}, children: sectionChildren });
      }

      // Generate DOCX file
      const buffer = await Packer.toBuffer(doc);
      const blob = new Blob(
        [buffer],
        { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
      );
      return {
        blob,
        filename: 'comprovacao_videos.docx'
      };
    } catch (error) {
      console.error('Error generating DOCX:', error);
      // Fallback to HTML if something goes wrong
      return this.generateHtmlDocument(videos, screenshots);
    }
  }

  // Helper function to get image data from URL
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
        resolve(canvas.toDataURL('image/png'));
      };
      img.onerror = reject;
      img.src = url;
    });
  }

  // Fallback: Generate HTML document
  generateHtmlDocument(videos, screenshots) {
    let html = `
      <!DOCTYPE html>
      <html>
      <head><meta charset="UTF-8"><title>Comprovação de Dados</title></head>
      <body>
        <h1>Comprovação de Dados - Playlist YouTube</h1>
    `;

    videos.forEach((video, i) => {
      html += `<h2>Vídeo ${i + 1}: ${video.title}</h2>`;
      html += `<ul>`;
      html += `<li><strong>Duração:</strong> ${video.duration}</li>`;
      html += `<li><strong>Views:</strong> ${video.views}</li>`;
      html += `<li><strong>Likes:</strong> ${video.likes}</li>`;
      html += `<li><strong>Link:</strong> <a href="https://www.youtube.com/watch?v=${video.videoId}">Ver no YouTube</a></li>`;
      html += `<li><strong>Data de Publicação:</strong> ${new Date(video.publishedDate).toLocaleDateString()}</li>`;
      html += `</ul>`;
      if (screenshots[i]) {
        html += `<img src="${screenshots[i].imageUrl}" style="max-width:600px;"><br>`;
      }
      html += `<hr>`;
    });

    html += `
      </body>
      </html>
    `;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    return {
      blob,
      filename: 'comprovacao_videos.html'
    };
  }
}
