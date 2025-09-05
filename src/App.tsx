import React, { useState, useCallback, useEffect, type ChangeEvent, type JSX, useRef } from 'react';
import { Upload, Presentation} from 'lucide-react';
import * as XLSX from 'xlsx';
import { renderAsync } from 'docx-preview';

type FileType = 'word' | 'excel' | 'powerpoint' | '';

interface ExcelSheetData {
  [sheetName: string]: (string | number)[][];
}

interface NativeMessage {
  type: 'READY' | 'LOADED' | 'WEB_ERROR';
  fileName?: string;
  message?: string;
}

interface IncomingMessage {
  type: 'LOAD_DOCUMENT';
  fileData: string;
  fileName: string;
}

declare global {
  interface Window {
    ReactNativeWebView?: {
      postMessage: (message: string) => void;
    };
  }
}

const DocumentViewer: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [fileType, setFileType] = useState<FileType>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [isEmbedded, setIsEmbedded] = useState<boolean>(false);
  const [debugInfo, setDebugInfo] = useState<string>('');
  const viewerRef = useRef<HTMLDivElement | null>(null);



  const postMessageToNative = (message: NativeMessage) => {
    console.log('Posting message to native:', message);
    if (window.ReactNativeWebView) {
      window.ReactNativeWebView.postMessage(JSON.stringify(message));
    }
  };

  const postErrorToNative = (message: string) => {
    postMessageToNative({ type: 'WEB_ERROR', message });
  };

  const addDebugInfo = (info: string) => {
    console.log('DEBUG:', info);
    setDebugInfo(prev => prev + '\n' + new Date().toLocaleTimeString() + ': ' + info);
  };

  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const viewMode = urlParams.get('viewMode');

    if (viewMode === 'embedded') {
      setIsEmbedded(true);
      addDebugInfo('Running in embedded mode');
      
      setTimeout(() => {
        postMessageToNative({ type: 'READY' });
        addDebugInfo('Sent READY signal to React Native');
      }, 500);
    }
  }, []);

  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      // Filter out docx-preview library errors
      if (typeof event.data === 'string' && 
          (event.data.includes('setImmediate') || event.data.includes('setImmedia'))) {
        return;
      }
      
      addDebugInfo(`Received message: ${event.data}`);
      
      try {
        const data: IncomingMessage = JSON.parse(event.data);
        addDebugInfo(`Parsed message type: ${data.type}`);
        
        if (data.type === 'LOAD_DOCUMENT') {
          addDebugInfo(`Loading document: ${data.fileName}, data length: ${data.fileData.length}`);
          loadFileFromBase64(data.fileData, data.fileName);
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error parsing message';
        
        // Filter out docx-preview library errors
        if (errorMessage.includes('setImmediate') || 
            errorMessage.includes('setImmedia') || 
            errorMessage.includes('not valid JSON')) {
          return;
        }
        
        addDebugInfo(`Error parsing message: ${errorMessage}`);
        postErrorToNative(errorMessage);
      }
    };

    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadFileFromBase64 = async (base64Data: string, fileName: string) => {
    setIsLoading(true);
    setError('');
    setExcelData(null);
    addDebugInfo(`Starting loadFileFromBase64 for ${fileName}`);
    
    try {
      addDebugInfo('Decoding base64...');
      const byteCharacters = atob(base64Data);
      addDebugInfo(`Base64 decoded: ${byteCharacters.length} bytes`);
      
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray]);
      addDebugInfo(`Blob created: ${blob.size} bytes`);
      
      const file = new File([blob], fileName);
      setSelectedFile(file);
      addDebugInfo('File object created, starting processFile...');
      
      await processFile(file);
      addDebugInfo('File processed successfully');
      
      postMessageToNative({ type: 'LOADED', fileName });
      
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Error loading file from data';
      addDebugInfo(`Error: ${errorMessage}`);
      setError(errorMessage);
      postErrorToNative(errorMessage);
    } finally {
      setIsLoading(false);
    }
  };

  const processFile = async (file: File) => {
    addDebugInfo(`Processing file: ${file.name}, size: ${file.size}`);
    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    addDebugInfo(`File extension: ${fileExtension}`);
    
    try {
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
        throw new Error('Unsupported file type');
      }
    } catch (error) {
      addDebugInfo(`Error in processFile: ${error}`);
      throw error;
    }
  };

  const processWordDocument = async (file: File): Promise<void> => {
    addDebugInfo('Processing Word document...');
    setFileType('word');
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      addDebugInfo(`ArrayBuffer size: ${arrayBuffer.byteLength}`);
      
      if (viewerRef.current) {
        addDebugInfo('Clearing viewer and rendering...');
        viewerRef.current.innerHTML = '';
        await renderAsync(arrayBuffer, viewerRef.current);
        addDebugInfo('Word document rendered successfully!');
      } else {
        throw new Error('Viewer container not available');
      }
    } catch (error) {
      addDebugInfo(`Error processing Word document: ${error}`);
      throw error;
    }
  };

  const processExcelDocument = async (file: File): Promise<void> => {
    addDebugInfo('Processing Excel document...');
    setFileType('excel');
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      
      const sheetsData: ExcelSheetData = {};
      workbook.SheetNames.forEach((sheetName: string) => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as (string | number)[][];
        sheetsData[sheetName] = jsonData;
      });
      
      setExcelData(sheetsData);
      addDebugInfo('Excel document processed successfully');
    } catch (error) {
      addDebugInfo(`Error processing Excel document: ${error}`);
      throw error;
    }
  };

  const processPowerPointDocument = (): void => {
    addDebugInfo('Processing PowerPoint document...');
    setFileType('powerpoint');
  };

  const handleFileUpload = useCallback(async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) return;

    setSelectedFile(file);
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
    <div key={sheetName} className="mb-4 sm:mb-8">
      <h3 className="text-base sm:text-lg font-semibold mb-2 sm:mb-4 text-gray-800 px-2">{sheetName}</h3>
      <div className="overflow-auto max-h-64 sm:max-h-96 border rounded-lg mx-2">
        <table className="min-w-full bg-white text-xs sm:text-sm">
          <tbody>
            {sheetData.map((row, rowIndex) => (
              <tr key={rowIndex} className={rowIndex === 0 ? 'bg-gray-50 font-medium' : 'hover:bg-gray-50'}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex} className="px-2 sm:px-4 py-1 sm:py-2 border-r border-b border-gray-200">
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

  // Debug Panel for embedded mode - Mobile optimized
  if (isEmbedded && (!selectedFile || error)) {
    return (
      <div className="p-2 sm:p-4 bg-white min-h-screen text-sm">
        <h2 className="text-lg font-bold mb-4">Document Viewer</h2>
        
        {isLoading && (
          <div className="mb-4 flex items-center">
            <div className="inline-block animate-spin rounded-full h-4 w-4 sm:h-6 sm:w-6 border-b-2 border-blue-600"></div>
            <span className="ml-2 text-sm">Processing...</span>
          </div>
        )}
        
        {error && (
          <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded text-sm">
            <strong>Error:</strong> {error}
          </div>
        )}
        
        <div className="bg-gray-50 p-3 rounded">
          <h3 className="font-semibold mb-2 text-sm">Debug Log:</h3>
          <pre className="text-xs whitespace-pre-wrap overflow-auto max-h-40">
            {debugInfo || 'Waiting for document...'}
          </pre>
        </div>
      </div>
    );
  }

  return (
    <div className={isEmbedded ? "bg-white min-h-screen" : "min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50"}>
      {/* Add viewport meta tag styles for mobile */}
      <style>{`
      .docx-container {
        width: 100% !important;
        max-width: 100% !important;
      }
      
      /* Mobile-first responsive styles for docx-preview */
      @media (max-width: 768px) {
        /* Override docx-preview page layout */
        .docx-wrapper {
          width: 100% !important;
          max-width: 100% !important;
          margin: 0 !important;
          padding: 8px !important;
        }
        
        /* Make pages responsive */
        .docx-wrapper > section {
          width: 100% !important;
          max-width: 100% !important;
          min-width: unset !important;
          margin: 0 0 16px 0 !important;
          padding: 12px !important;
          box-shadow: none !important;
          border: 1px solid #e5e5e5 !important;
        }
        
        /* Responsive text */
        .docx-wrapper p {
          font-size: 14px !important;
          line-height: 1.4 !important;
          margin: 8px 0 !important;
          word-wrap: break-word !important;
          overflow-wrap: break-word !important;
        }
        
        /* Responsive headings */
        .docx-wrapper h1 { font-size: 20px !important; margin: 12px 0 8px 0 !important; }
        .docx-wrapper h2 { font-size: 18px !important; margin: 10px 0 6px 0 !important; }
        .docx-wrapper h3 { font-size: 16px !important; margin: 8px 0 4px 0 !important; }
        .docx-wrapper h4, .docx-wrapper h5, .docx-wrapper h6 { 
          font-size: 14px !important; 
          margin: 6px 0 4px 0 !important; 
        }
        
        /* Responsive tables */
        .docx-wrapper table {
          width: 100% !important;
          max-width: 100% !important;
          font-size: 12px !important;
          display: block !important;
          overflow-x: auto !important;
          white-space: nowrap !important;
        }
        
        .docx-wrapper table thead,
        .docx-wrapper table tbody,
        .docx-wrapper table tr {
          display: table !important;
          width: 100% !important;
        }
        
        .docx-wrapper table td,
        .docx-wrapper table th {
          padding: 4px 6px !important;
          font-size: 11px !important;
          border: 1px solid #ddd !important;
        }
        
        /* Responsive images */
        .docx-wrapper img {
          max-width: 100% !important;
          height: auto !important;
          display: block !important;
          margin: 8px auto !important;
        }
        
        /* Responsive lists */
        .docx-wrapper ul, .docx-wrapper ol {
          padding-left: 20px !important;
          margin: 8px 0 !important;
        }
        
        .docx-wrapper li {
          margin: 4px 0 !important;
          font-size: 14px !important;
          line-height: 1.4 !important;
        }
        
        /* Remove fixed widths */
        .docx-wrapper * {
          max-width: 100% !important;
        }
        
        /* Override any absolute positioning */
        .docx-wrapper [style*="position: absolute"] {
          position: relative !important;
        }
        
        /* Force page breaks to be ignored on mobile */
        .docx-wrapper .docx-page-break {
          display: none !important;
        }
        
        /* Text selection and zoom */
        .docx-wrapper {
          -webkit-text-size-adjust: 100% !important;
          -ms-text-size-adjust: 100% !important;
          text-size-adjust: 100% !important;
        }
      }
      
      /* Tablet styles */
      @media (min-width: 769px) and (max-width: 1024px) {
        .docx-wrapper > section {
          max-width: 95% !important;
          margin: 0 auto 20px auto !important;
        }
        
        .docx-wrapper p {
          font-size: 15px !important;
        }
      }
    `}</style>
      
      <div className={isEmbedded ? "p-1 sm:p-2" : "container mx-auto px-4 py-8"}>
        
        {!isEmbedded && (
          <div className="text-center mb-8">
            <h1 className="text-2xl sm:text-4xl font-bold text-gray-800 mb-2">Document Viewer</h1>
            <p className="text-gray-600 text-sm sm:text-base">Upload and view Word, Excel, and PowerPoint files</p>
          </div>
        )}

        {!isEmbedded && !selectedFile && (
          <div className="max-w-2xl mx-auto mb-8">
            <div className="bg-white rounded-xl shadow-lg p-4 sm:p-6 border-2 border-dashed border-gray-300 hover:border-blue-400 transition-colors">
              <div className="text-center">
                <Upload className="w-8 h-8 sm:w-12 sm:h-12 text-gray-400 mx-auto mb-4" />
                <label htmlFor="file-upload" className="cursor-pointer">
                  <span className="text-base sm:text-lg font-medium text-gray-700 hover:text-blue-600">
                    Choose a document to upload
                  </span>
                  <input
                    id="file-upload"
                    type="file"
                    className="hidden"
                    accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx"
                    onChange={handleFileUpload}
                  />
                </label>
                <p className="text-xs sm:text-sm text-gray-500 mt-2">
                  Supports Word (.doc, .docx), Excel (.xls, .xlsx), and PowerPoint (.ppt, .pptx)
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Word Document Content - Mobile Optimized */}
        {fileType === 'word' && (
          <div className={isEmbedded ? "w-full" : "max-w-sm mx-auto"}>
            <div className={isEmbedded ? "w-full" : "bg-white rounded-xl shadow-lg p-4 sm:p-8"}>
              <div 
                ref={viewerRef}
                className={` ${isEmbedded ? 'min-h-screen' : 'min-h-96'} overflow-auto docx-preview`}
                style={{ 
                  backgroundColor: 'white',
                  fontSize: isEmbedded ? '12px' : '11px',
                  lineHeight: '1.4',
                
                }}
              />
            </div>
          </div>
        )}

        {/* PowerPoint Content - Mobile Optimized */}
        {fileType === 'powerpoint' && (
          <div className="text-center py-8 sm:py-16 px-4">
            <Presentation className="w-12 h-12 sm:w-16 sm:h-16 text-gray-400 mx-auto mb-4" />
            <h3 className="text-lg sm:text-xl font-semibold text-gray-800 mb-2">PowerPoint Preview</h3>
            <p className="text-gray-600 text-sm sm:text-base">PowerPoint files are loaded but preview is not yet implemented.</p>
          </div>
        )}

        {/* Excel Content - Mobile Optimized */}
        {excelData && (
          <div className={isEmbedded ? "w-full" : "max-w-6xl mx-auto"}>
            <div className={isEmbedded ? "" : "bg-white rounded-xl shadow-lg p-4 sm:p-8"}>
              {!isEmbedded && <h2 className="text-xl sm:text-2xl font-bold text-gray-800 mb-4 sm:mb-6">Excel Workbook</h2>}
              {Object.entries(excelData).map(([sheetName, sheetData]) => 
                renderExcelSheet(sheetData, sheetName)
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default DocumentViewer;
