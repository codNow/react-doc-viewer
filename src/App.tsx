import React, { useState, useCallback, useEffect, type ChangeEvent, type JSX, useRef } from 'react';
import { Upload, Presentation } from 'lucide-react';
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
  const [loadSource, setLoadSource] = useState<'upload' | 'url' | 'mobile'>('upload');
  const [isEmbedded, setIsEmbedded] = useState<boolean>(false);
  const [debugInfo, setDebugInfo] = useState<string>('');
  const viewerRef = useRef<HTMLDivElement | null>(null);



  const postMessageToNative = (message: NativeMessage) => {
    console.log('Posting message to native:', message);
    if (window.ReactNativeWebView) {
      window.ReactNativeWebView.postMessage(JSON.stringify(message));
    } else {
      console.log('ReactNativeWebView not available - probably in browser');
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
      
      // Send ready signal
      setTimeout(() => {
        postMessageToNative({ type: 'READY' });
        addDebugInfo('Sent READY signal to React Native');
      }, 500);
    }
  }, []);

  // In your DocumentViewer.tsx, update the handleMessage function:

useEffect(() => {
  const handleMessage = (event: MessageEvent) => {
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
      if (errorMessage.includes('setImmediate') || errorMessage.includes('setImmedia')) {
        // Ignore these errors - they're from the docx-preview library
        console.warn('Ignoring docx-preview library error:', errorMessage);
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

  // Debug Panel for embedded mode
  if (isEmbedded && (!selectedFile || error)) {
    return (
      <div className="p-4 bg-white min-h-screen">
        <h2 className="text-lg font-bold mb-4">Document Viewer Debug</h2>
        
        {isLoading && (
          <div className="mb-4">
            <div className="inline-block animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
            <span className="ml-2">Processing document...</span>
            <div className="mb-2">
            <strong>Load Source:</strong> {loadSource}
          </div>

          </div>
        )}
        
        {error && (
          <div className="mb-4 p-4 bg-red-50 border border-red-200 rounded">
            <strong>Error:</strong> {error}
          </div>
        )}
        
        <div className="bg-gray-50 p-4 rounded">
          <h3 className="font-semibold mb-2">Debug Log:</h3>
          <pre className="text-xs whitespace-pre-wrap">{debugInfo || 'No debug info yet...'}</pre>
        </div>
      </div>
    );
  }

  return (
    <div className={isEmbedded ? "bg-white min-h-screen" : "min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50"}>
      <div className={isEmbedded ? "px-2 py-2" : "container mx-auto px-4 py-8"}>
        
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
                    id="file-upload"
                    type="file"
                    className="hidden"
                    accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx"
                    onChange={handleFileUpload}
                  />
                </label>
                <p className="text-sm text-gray-500 mt-2">
                  Supports Word (.doc, .docx), Excel (.xls, .xlsx), and PowerPoint (.ppt, .pptx)
                </p>
              </div>
            </div>
          </div>
        )}

        {fileType === 'word' && (
          <div className={isEmbedded ? "" : "max-w-4xl mx-auto"}>
            <div className={isEmbedded ? "" : "bg-white rounded-xl shadow-lg p-8"}>
              <div 
                ref={viewerRef}
                className={`w-full ${isEmbedded ? 'min-h-screen' : 'min-h-96'} overflow-auto`}
                style={{ backgroundColor: 'white' }}
              />
            </div>
          </div>
        )}

        {fileType === 'powerpoint' && (
          <div className="text-center py-16">
            <Presentation className="w-16 h-16 text-gray-400 mx-auto mb-4" />
            <h3 className="text-xl font-semibold text-gray-800 mb-2">PowerPoint Preview</h3>
            <p className="text-gray-600">PowerPoint files are loaded but preview is not yet implemented.</p>
          </div>
        )}

        {excelData && (
          <div className={isEmbedded ? "" : "max-w-6xl mx-auto"}>
            <div className={isEmbedded ? "" : "bg-white rounded-xl shadow-lg p-8"}>
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
