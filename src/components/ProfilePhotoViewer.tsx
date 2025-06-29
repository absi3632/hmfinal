import React from 'react';
import { X, Download } from 'lucide-react';

interface ProfilePhotoViewerProps {
  photo: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
  };
  housemaidName: string;
  onClose: () => void;
}

const ProfilePhotoViewer: React.FC<ProfilePhotoViewerProps> = ({ photo, housemaidName, onClose }) => {
  if (!photo.fileData || !photo.fileName) return null;

  const handleDownload = () => {
    try {
      const link = document.createElement('a');
      link.href = photo.fileData!;
      link.download = photo.fileName!;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading photo. Please try again.');
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg max-w-2xl w-full max-h-[90vh] overflow-hidden">
        <div className="flex items-center justify-between p-6 border-b">
          <div>
            <h3 className="text-xl font-semibold text-gray-900">Profile Photo</h3>
            <p className="text-sm text-gray-600 mt-1">{housemaidName}</p>
          </div>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
              title="Download Photo"
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
        <div className="p-6 flex justify-center">
          <img
            src={photo.fileData}
            alt={`Profile photo of ${housemaidName}`}
            className="max-w-full max-h-[60vh] object-contain rounded-lg shadow-lg"
          />
        </div>
        <div className="px-6 pb-6">
          <div className="text-sm text-gray-500 text-center">
            <p>{photo.fileName}</p>
            {photo.uploadDate && (
              <p>Uploaded: {new Date(photo.uploadDate).toLocaleDateString()}</p>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default ProfilePhotoViewer;