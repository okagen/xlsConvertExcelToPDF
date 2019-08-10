# xlsConvertExcelToPDF
Convert some sheets in Excel files to PDF files.   
複数のExcelファイル内のシートをPDFファイルに変換する。

## Usage

### step_1 Click the "リスト作成" button.
Search in the directry which contains the xlsConvertToPDF.xlsm and list Excel files from all folders and sub-folders.  
xlsConvertToPDF.xlsmを含むディレクトリ内を検索し、処理対象のExcelファイルのリストをシートに書き出し。

Image of Tool Sheet after create the list.<br>
<img src="https://github.com/okagen/xlsConvertExcelToPDF/blob/master/Data/tool01.png?raw=true" width="600">

### step_2 Click the "PDF作成" button.
Each sheets of the Excel files are converted to PDF and output to the specified output directory.  
処理対処Excelファイルの各シートがPDFに変換され、指定された出力先に出力される。

Image of Tool Sheet after output PDF files.<br>
<img src="https://github.com/okagen/xlsConvertExcelToPDF/blob/master/Data/tool02.png?raw=true" width="600">

## Example

The Excel file contains 3 sheets as below. Each sheets has a list with header, print area and page number.  
Excelファイルに以下の３シート含まれ、その中にヘッダー付きのリストがあり、印刷範囲、ページ番号が設定されています。

  - Sheet01 will be one page. Sheet01は1ページ分
  - Sheet02 will be two pages. Sheet02は2ページ分
  - Sheet03 will be three pages. Sheet03には3ページ分
 
Image of the Excel file.<br>
<img src="https://github.com/okagen/xlsConvertExcelToPDF/blob/master/Data/Book01_xlsx.png?raw=true" width="200">


After convert the Excel file to PDF file, the PDF file will contain six pages in it.  
ExcelファイルからにコンバートされたPDFファイルには、６ページ分が含まれます。

Image of the PDF file exported.<br>
<img src="https://github.com/okagen/xlsConvertExcelToPDF/blob/master/Data/Book01_pdf.png?raw=true" width="200">




