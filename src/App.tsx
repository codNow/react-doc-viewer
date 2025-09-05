import React, { useState, useCallback, useEffect, type ChangeEvent, type JSX, useRef } from 'react';
import { Upload, FileText, FileSpreadsheet, Presentation, AlertCircle, Eye, Smartphone, Globe } from 'lucide-react';
import * as XLSX from 'xlsx';
import { renderAsync } from 'docx-preview';

type FileType = 'word' | 'excel' | 'powerpoint' | '';

interface ExcelSheetData {
  [sheetName: string]: (string | number)[][];
}

const DocumentViewer: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [fileType, setFileType] = useState<FileType>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [loadSource, setLoadSource] = useState<'upload' | 'url' | 'mobile'>('upload');
  const [isEmbedded, setIsEmbedded] = useState<boolean>(false); // State for embedded mode
  const viewerRef = useRef<HTMLDivElement | null>(null);

  const getFileIcon = (type: FileType): JSX.Element => {
    if (type.includes('word') || type.includes('document')) return <FileText className="w-6 h-6" />;
    if (type.includes('sheet') || type.includes('excel')) return <FileSpreadsheet className="w-6 h-6" />;
    if (type.includes('presentation') || type.includes('powerpoint')) return <Presentation className="w-6 h-6" />;
    return <FileText className="w-6 h-6" />;
  };

  // Check URL parameters on component mount
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');
    const fileName = urlParams.get('fileName');
    const viewMode = urlParams.get('viewMode');

    if (viewMode === 'embedded') {
      setIsEmbedded(true);
    }
    
    if (fileUrl && fileName) {
      setLoadSource('url');
      loadFileFromUrl(fileUrl, fileName);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Listen for postMessage from React Native WebView
  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      try {
        const data = JSON.parse(event.data);
        if (data.type === 'LOAD_DOCUMENT') {
          setLoadSource('mobile');
          const { fileData, fileName } = data;
          loadFileFromBase64(fileData, fileName);
        }
      } catch (error) {
        console.error('Error parsing message:', error);
      }
    };

    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadFileFromUrl = async (url: string, fileName: string) => {
    setIsLoading(true);
    setError('');
    setExcelData(null);
    
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error('Failed to fetch file');
      
      const arrayBuffer = await response.arrayBuffer();
      const file = new File([arrayBuffer], fileName, {
        type: response.headers.get('content-type') || 'application/octet-stream'
      });
      
      setSelectedFile(file);
      await processFile(file);
      
      if (window.parent !== window) {
        window.parent.postMessage({ type: 'LOADED', fileName }, '*');
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Error loading file from URL';
      setError(errorMessage);
      
      if (window.parent !== window) {
        window.parent.postMessage({ type: 'ERROR', message: errorMessage }, '*');
      }
    } finally {
      setIsLoading(false);
    }
  };

  const loadFileFromBase64 = async (base64Data: string, fileName: string) => {
    setIsLoading(true);
    setError('');
    setExcelData(null);
    
    try {
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray]);
      
      const file = new File([blob], fileName);
      setSelectedFile(file);
      await processFile(file);
      
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Error loading file from data';
      setError(errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const processFile = async (file: File) => {
    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    
    if (fileExtension === 'docx' || fileExtension === 'doc') {
      await processWordDocument(file);
    } 
    else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      await processExcelDocument(file);
    }
    else if (fileExtension === 'pptx' || fileExtension === 'ppt') {
      processPowerPointDocument();
    }
    else {
      throw new Error('Unsupported file type. Please upload Word (.docx, .doc), Excel (.xlsx, .xls), or PowerPoint (.pptx, .ppt) files.');
    }
  };

  const processWordDocument = async (file: File): Promise<void> => {
    setFileType('word');
    if (!file) return;
    
    const arrayBuffer = await file.arrayBuffer();
    if (viewerRef.current) {
      viewerRef.current.innerHTML = '';
      await renderAsync(arrayBuffer, viewerRef.current);
    }
  };

  const processExcelDocument = async (file: File): Promise<void> => {
    setFileType('excel');
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    const sheetsData: ExcelSheetData = {};
    workbook.SheetNames.forEach((sheetName: string) => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as (string | number)[][];
      sheetsData[sheetName] = jsonData;
    });
    
    setExcelData(sheetsData);
  };

  const processPowerPointDocument = (): void => {
    setFileType('powerpoint');
  };

  const handleFileUpload = useCallback(async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) return;

    setSelectedFile(file);
    setLoadSource('upload');
    setIsLoading(true);
    setError('');
    setExcelData(null);

    try {
      await processFile(file);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Error processing file';
      setError(errorMessage);
    } finally {
      setIsLoading(false);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const renderExcelSheet = (sheetData: (string | number)[][], sheetName: string): JSX.Element => (
    <div key={sheetName} className="mb-8">
      <h3 className="text-lg font-semibold mb-4 text-gray-800">{sheetName}</h3>
      <div className="overflow-auto max-h-96 border rounded-lg">
        <table className="min-w-full bg-white">
          <tbody>
            {sheetData.map((row, rowIndex) => (
              <tr key={rowIndex} className={rowIndex === 0 ? 'bg-gray-50 font-medium' : 'hover:bg-gray-50'}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex} className="px-4 py-2 border-r border-b border-gray-200 text-sm">
                    {cell || ''}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );

  const formatFileSize = (bytes: number): string => (bytes / 1024 / 1024).toFixed(2);
  const capitalizeFileType = (type: FileType): string => type.charAt(0).toUpperCase() + type.slice(1);
  const getLoadSourceIcon = () => {
    switch (loadSource) {
      case 'mobile': return <Smartphone className="w-5 h-5 mr-1 text-blue-600" />;
      case 'url': return <Globe className="w-5 h-5 mr-1 text-green-600" />;
      default: return <Upload className="w-5 h-5 mr-1 text-purple-600" />;
    }
  };
  const getLoadSourceText = () => {
    switch (loadSource) {
      case 'mobile': return 'Loaded from Mobile App';
      case 'url': return 'Loaded from URL';
      default: return 'Uploaded from Device';
    }
  };

  return (
    <div className={isEmbedded ? "bg-white" : "min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50"}>
      <div className={isEmbedded ? "" : "container mx-auto px-4 py-8"}>
        
        {!isEmbedded && (
          <div className="text-center mb-8">
            <h1 className="text-4xl font-bold text-gray-800 mb-2">Document Viewer</h1>
            <p className="text-gray-600">Upload and view Word, Excel, and PowerPoint files</p>
          </div>
        )}

        {!isEmbedded && !selectedFile && (
          <div className="max-w-2xl mx-auto mb-8">
            <div className="bg-white rounded-xl shadow-lg p-6 border-2 border-dashed border-gray-300 hover:border-blue-400 transition-colors">
              <div className="text-center">
                <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
                <label htmlFor="file-upload" className="cursor-pointer">
                  <span className="text-lg font-medium text-gray-700 hover:text-blue-600">
                    Choose a document to upload
                  </span>
                  <input
                    id="file-upload" type="file" className="hidden"
                    accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx"
                    onChange={handleFileUpload}
                  />
                </label>
                <p className="text-sm text-gray-500 mt-2">
                  Supports Word (.doc, .docx), Excel (.xls, .xlsx), and PowerPoint (.ppt, .pptx)
                </p>
                <p className="text-xs text-gray-400 mt-3">
                  💡 This viewer can also be used from mobile apps via WebView integration
                </p>
              </div>
            </div>
          </div>
        )}

        {isLoading && (
          <div className="text-center py-8">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
            <p className="mt-2 text-gray-600">Processing document...</p>
          </div>
        )}

        {error && (
          <div className="max-w-2xl mx-auto mb-8">
            <div className="bg-red-50 border border-red-200 rounded-lg p-4 flex items-center">
              <AlertCircle className="w-5 h-5 text-red-500 mr-3" />
              <span className="text-red-700">{error}</span>
            </div>
          </div>
        )}

        {selectedFile && !isLoading && !isEmbedded && (
          <div className="max-w-4xl mx-auto mb-6">
            <div className="bg-white rounded-lg shadow-md p-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center">
                  {getFileIcon(fileType)}
                  <div className="ml-3">
                    <h3 className="font-medium text-gray-800">{selectedFile.name}</h3>
                    <p className="text-sm text-gray-500">
                      {formatFileSize(selectedFile.size)} MB • {capitalizeFileType(fileType)}
                    </p>
                  </div>
                </div>
                <div className="flex flex-col items-end">
                  <div className="flex items-center text-green-600 mb-1">
                    <Eye className="w-5 h-5 mr-1" />
                    <span className="text-sm font-medium">Viewing</span>
                  </div>
                  <div className="flex items-center text-gray-500">
                    {getLoadSourceIcon()}
                    <span className="text-xs">{getLoadSourceText()}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {fileType === 'word' && (
          <div className={isEmbedded ? "" : "max-w-4xl mx-auto"}>
            <div className={isEmbedded ? "" : "bg-white rounded-xl shadow-lg p-8"}>
              <div 
                ref={viewerRef}
                className="w-full min-h-screen" // Use min-h-screen in embedded mode
                style={{ backgroundColor: 'white' }}
              />
            </div>
          </div>
        )}

        {fileType === 'powerpoint' && (
          <div className="max-w-4xl mx-auto">
            <div className="bg-white rounded-xl shadow-lg p-8">
              <div className="text-center py-16">
                <Presentation className="w-16 h-16 text-gray-400 mx-auto mb-4" />
                <h3 className="text-xl font-semibold text-gray-800 mb-2">PowerPoint Preview</h3>
                <p className="text-gray-600 mb-4">
                  PowerPoint files are loaded but preview is not yet implemented.
                </p>
                <p className="text-sm text-gray-500">
                  File: {selectedFile?.name}
                </p>
              </div>
            </div>
          </div>
        )}

        {excelData && (
          <div className={isEmbedded ? "" : "max-w-6xl mx-auto"}>
            <div className={isEmbedded ? "" : "bg-white rounded-xl shadow-lg p-8"}>
              {!isEmbedded && <h2 className="text-2xl font-bold text-gray-800 mb-6">Excel Workbook</h2>}
              {Object.entries(excelData).map(([sheetName, sheetData]) => 
                renderExcelSheet(sheetData, sheetName)
              )}
            </div>
          </div>
        )}

        {!isEmbedded && !selectedFile && (
          <div className="max-w-4xl mx-auto mt-16">
            <div className="grid md:grid-cols-3 gap-8">
              <div className="text-center">
                <div className="bg-blue-100 rounded-full w-16 h-16 flex items-center justify-center mx-auto mb-4">
                  <FileText className="w-8 h-8 text-blue-600" />
                </div>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Word Documents</h3>
                <p className="text-gray-600">View .doc and .docx files with full formatting preserved</p>
              </div>
              <div className="text-center">
                <div className="bg-green-100 rounded-full w-16 h-16 flex items-center justify-center mx-auto mb-4">
                  <FileSpreadsheet className="w-8 h-8 text-green-600" />
                </div>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Excel Spreadsheets</h3>
                <p className="text-gray-600">Display .xls and .xlsx files with all worksheets</p>
              </div>
              <div className="text-center">
                <div className="bg-purple-100 rounded-full w-16 h-16 flex items-center justify-center mx-auto mb-4">
                  <Presentation className="w-8 h-8 text-purple-600" />
                </div>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">PowerPoint</h3>
                <p className="text-gray-600">Preview .ppt and .pptx presentations</p>
              </div>
            </div>
            <div className="mt-16 bg-gradient-to-r from-blue-50 to-purple-50 rounded-xl p-8">
              <h3 className="text-2xl font-bold text-gray-800 mb-4 text-center">Mobile Integration</h3>
              <div className="grid md:grid-cols-2 gap-8">
                <div>
                  <h4 className="font-semibold text-gray-800 mb-2 flex items-center">
                    <Globe className="w-5 h-5 mr-2 text-green-600" />
                    URL Parameters
                  </h4>
                  <p className="text-gray-600 text-sm">
                    Load documents via URL: <code className="bg-gray-200 px-2 py-1 rounded">?fileUrl=...&fileName=...</code>
                  </p>
                </div>
                <div>
                  <h4 className="font-semibold text-gray-800 mb-2 flex items-center">
                    <Smartphone className="w-5 h-5 mr-2 text-blue-600" />
                    WebView Integration
                  </h4>
                  <p className="text-gray-600 text-sm">
                    Accepts documents via postMessage from React Native WebView
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default DocumentViewer;
