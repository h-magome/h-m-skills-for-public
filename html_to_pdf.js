const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

async function convertHtmlToPdf(htmlFilePath, pdfFilePath) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    // HTMLファイルを読み込み
    const htmlContent = fs.readFileSync(htmlFilePath, 'utf8');
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
    
    // PDFオプション
    const pdfOptions = {
        path: pdfFilePath,
        format: 'A4',
        margin: {
            top: '20mm',
            right: '20mm',
            bottom: '20mm',
            left: '20mm'
        },
        printBackground: true,
        displayHeaderFooter: true,
        headerTemplate: '<div style="font-size: 10px; width: 100%; text-align: center; color: #666;">HM スキルシート</div>',
        footerTemplate: '<div style="font-size: 10px; width: 100%; text-align: center; color: #666;">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>',
        preferCSSPageSize: true
    };
    
    // PDFを生成
    await page.pdf(pdfOptions);
    await browser.close();
    
    console.log(`✓ PDFファイルが作成されました: ${pdfFilePath}`);
}

async function main() {
    const htmlFiles = [
        'HM_スキルシート.html',
        'README.html'
    ];
    
    console.log('HTMLからPDFへの変換を開始します...\n');
    
    for (const htmlFile of htmlFiles) {
        if (fs.existsSync(htmlFile)) {
            const pdfFile = htmlFile.replace('.html', '.pdf');
            try {
                await convertHtmlToPdf(htmlFile, pdfFile);
            } catch (error) {
                console.error(`✗ エラー: ${htmlFile} の変換に失敗しました - ${error.message}`);
            }
        } else {
            console.log(`✗ HTMLファイルが見つかりません: ${htmlFile}`);
        }
    }
    
    console.log('\n変換処理が完了しました！');
}

main().catch(console.error); 