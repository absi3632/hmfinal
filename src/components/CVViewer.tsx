import React from 'react';
import { X, Download, File } from 'lucide-react';

interface CVViewerProps {
  cv: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
  };
  onClose: () => void;
}

const CVViewer: React.FC<CVViewerProps> = ({ cv, onClose }) => {
  if (!cv.fileData || !cv.fileName) return null;

  const handleDownload = () => {
    try {
      const link = document.createElement('a');
      link.href = cv.fileData!;
      link.download = cv.fileName!;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading file. Please try again.');
    }
  };

  const renderContent = () => {
    if (!cv.fileType || !cv.fileData) {
      return (
        <div className="flex items-center justify-center h-96">
          <div className="text-center">
            <File className="h-16 w-16 text-gray-400 mx-auto mb-4" />
            <p className="text-gray-600">Unable to preview this file type</p>
            <button
              onClick={handleDownload}
              className="mt-4 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors"
            >
              Download to View
            </button>
          </div>
        </div>
      );
    }

    // Handle PDF files
    if (cv.fileType === 'application/pdf') {
      return (
        <iframe
          src={cv.fileData}
          className="w-full h-96 border rounded-lg"
          title={`CV Preview - ${cv.fileName}`}
        />
      );
    }

    // Handle image files
    if (cv.fileType.startsWith('image/')) {
      return (
        <div className="flex justify-center">
          <img
            src={cv.fileData}
            alt={`CV - ${cv.fileName}`}
            className="max-w-full max-h-96 object-contain border rounded-lg"
          />
        </div>
      );
    }

    // Handle Word documents and other files
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <File className="h-16 w-16 text-gray-400 mx-auto mb-4" />
          <p className="text-gray-600 mb-2">Preview not available for this file type</p>
          <p className="text-sm text-gray-500 mb-4">
            {cv.fileName} ({cv.fileType})
          </p>
          <button
            onClick={handleDownload}
            className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors flex items-center space-x-2 mx-auto"
          >
            <Download className="h-4 w-4" />
            <span>Download to View</span>
          </button>
        </div>
      </div>
    );
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg max-w-4xl w-full max-h-[90vh] overflow-hidden">
        <div className="flex items-center justify-between p-6 border-b">
          <div>
            <h3 className="text-xl font-semibold text-gray-900">CV Preview</h3>
            <p className="text-sm text-gray-600 mt-1">{cv.fileName}</p>
          </div>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
              title="Download CV"
            >
              <Download className="h-5 w-5" />
            </button>
            <button
              onClick={onClose}
              className="p-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-lg transition-colors"
              title="Close"
            >
              <X className="h-5 w-5" />
            </button>
          </div>
        </div>
        <div className="p-6 overflow-y-auto max-h-[calc(90vh-120px)]">
          {renderContent()}
        </div>
      </div>
    </div>
  );
};

export default CVViewer;