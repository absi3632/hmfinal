import React from 'react';
import { X, Download, Ticket, Plane } from 'lucide-react';

interface AirTicketViewerProps {
  airTicket: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
    ticketNumber?: string;
    bookingReference?: string;
  };
  housemaidName: string;
  onClose: () => void;
}

const AirTicketViewer: React.FC<AirTicketViewerProps> = ({ airTicket, housemaidName, onClose }) => {
  if (!airTicket.fileData || !airTicket.fileName) return null;

  const handleDownload = () => {
    try {
      const link = document.createElement('a');
      link.href = airTicket.fileData!;
      link.download = airTicket.fileName!;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading file. Please try again.');
    }
  };

  const renderContent = () => {
    if (!airTicket.fileType || !airTicket.fileData) {
      return (
        <div className="flex items-center justify-center h-96">
          <div className="text-center">
            <Ticket className="h-16 w-16 text-gray-400 mx-auto mb-4" />
            <p className="text-gray-600">Unable to preview this file type</p>
            <button
              onClick={handleDownload}
              className="mt-4 bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 transition-colors"
            >
              Download to View
            </button>
          </div>
        </div>
      );
    }

    // Handle PDF files
    if (airTicket.fileType === 'application/pdf') {
      return (
        <iframe
          src={airTicket.fileData}
          className="w-full h-96 border rounded-lg"
          title={`Air Ticket Preview - ${airTicket.fileName}`}
        />
      );
    }

    // Handle image files
    if (airTicket.fileType.startsWith('image/')) {
      return (
        <div className="flex justify-center">
          <img
            src={airTicket.fileData}
            alt={`Air Ticket - ${airTicket.fileName}`}
            className="max-w-full max-h-96 object-contain border rounded-lg"
          />
        </div>
      );
    }

    // Handle Word documents and other files
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <Ticket className="h-16 w-16 text-gray-400 mx-auto mb-4" />
          <p className="text-gray-600 mb-2">Preview not available for this file type</p>
          <p className="text-sm text-gray-500 mb-4">
            {airTicket.fileName} ({airTicket.fileType})
          </p>
          <button
            onClick={handleDownload}
            className="bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 transition-colors flex items-center space-x-2 mx-auto"
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
        <div className="flex items-center justify-between p-6 border-b bg-purple-50">
          <div className="flex items-center space-x-3">
            <div className="p-2 bg-purple-100 rounded-lg">
              <Plane className="h-6 w-6 text-purple-600" />
            </div>
            <div>
              <h3 className="text-xl font-semibold text-gray-900">Air Ticket</h3>
              <p className="text-sm text-gray-600 mt-1">{housemaidName}</p>
              <p className="text-sm text-purple-600 font-medium">{airTicket.fileName}</p>
              <div className="flex items-center space-x-4 mt-1">
                {airTicket.ticketNumber && (
                  <p className="text-sm text-purple-700">
                    Ticket: {airTicket.ticketNumber}
                  </p>
                )}
                {airTicket.bookingReference && (
                  <p className="text-sm text-purple-700">
                    Ref: {airTicket.bookingReference}
                  </p>
                )}
              </div>
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-purple-600 hover:bg-purple-50 rounded-lg transition-colors"
              title="Download Air Ticket"
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

export default AirTicketViewer;