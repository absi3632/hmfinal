import React from 'react';
import { X, Download, FileText, Shield } from 'lucide-react';

interface POLOClearanceViewerProps {
  clearance: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
    completionDate?: string;
  };
  housemaidName: string;
  onClose: () => void;
}

const POLOClearanceViewer: React.FC<POLOClearanceViewerProps> = ({ clearance, housemaidName, onClose }) => {
  if (!clearance.fileData || !clearance.fileName) return null;

  const handleDownload = () => {
    try {
      const link = document.createElement('a');
      link.href = clearance.fileData!;
      link.download = clearance.fileName!;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading file. Please try again.');
    }
  };

  const renderContent = () => {
    if (!clearance.fileType || !clearance.fileData) {
      return (
        <div className="flex items-center justify-center h-96">
          <div className="text-center">
            <FileText className="h-16 w-16 text-gray-400 mx-auto mb-4" />
            <p className="text-gray-600">Unable to preview this file type</p>
            <button
              onClick={handleDownload}
              className="mt-4 bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 transition-colors"
            >
              Download to View
            </button>
          </div>
        </div>
      );
    }

    // Handle PDF files
    if (clearance.fileType === 'application/pdf') {
      return (
        <iframe
          src={clearance.fileData}
          className="w-full h-96 border rounded-lg"
          title={`POLO Clearance Preview - ${clearance.fileName}`}
        />
      );
    }

    // Handle Word documents and other files
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <FileText className="h-16 w-16 text-gray-400 mx-auto mb-4" />
          <p className="text-gray-600 mb-2">Preview not available for this file type</p>
          <p className="text-sm text-gray-500 mb-4">
            {clearance.fileName} ({clearance.fileType})
          </p>
          <button
            onClick={handleDownload}
            className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 transition-colors flex items-center space-x-2 mx-auto"
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
        <div className="flex items-center justify-between p-6 border-b bg-green-50">
          <div className="flex items-center space-x-3">
            <div className="p-2 bg-green-100 rounded-lg">
              <Shield className="h-6 w-6 text-green-600" />
            </div>
            <div>
              <h3 className="text-xl font-semibold text-gray-900">POLO Clearance</h3>
              <p className="text-sm text-gray-600 mt-1">{housemaidName}</p>
              <p className="text-sm text-green-600 font-medium">{clearance.fileName}</p>
              {clearance.completionDate && (
                <p className="text-sm text-green-700">
                  Completed: {new Date(clearance.completionDate).toLocaleDateString()}
                </p>
              )}
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
              title="Download POLO Clearance"
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

export default POLOClearanceViewer;