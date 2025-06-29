import React, { useState, useRef } from 'react';
import { 
  Upload, 
  Save, 
  RotateCcw, 
  Image, 
  Palette, 
  Type, 
  X, 
  Eye, 
  Trash2,
  Download,
  Building,
  Copyright
} from 'lucide-react';
import { BrandSettings as BrandSettingsType } from '../types/brand';
import BrandLogo from './BrandLogo';

interface BrandSettingsProps {
  brandSettings: BrandSettingsType;
  onSave: (settings: BrandSettingsType) => void;
  onClose: () => void;
}

const BrandSettings: React.FC<BrandSettingsProps> = ({ brandSettings, onSave, onClose }) => {
  const [formData, setFormData] = useState<BrandSettingsType>(brandSettings);
  const [isUploading, setIsUploading] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleLogoUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type - only images
    const allowedTypes = [
      'image/jpeg',
      'image/png',
      'image/jpg',
      'image/webp',
      'image/svg+xml'
    ];

    if (!allowedTypes.includes(file.type)) {
      alert('Please upload an image file (JPG, PNG, WebP, SVG).');
      return;
    }

    // Validate file size (1MB limit for logos)
    if (file.size > 1 * 1024 * 1024) {
      alert('Logo size must be less than 1MB.');
      return;
    }

    setIsUploading(true);

    try {
      const reader = new FileReader();
      reader.onload = () => {
        const base64Data = reader.result as string;
        setFormData(prev => ({
          ...prev,
          logoFileName: file.name,
          logoFileData: base64Data,
          logoFileType: file.type,
          logoUploadDate: new Date().toISOString()
        }));
        setIsUploading(false);
      };
      reader.onerror = () => {
        alert('Error reading logo file. Please try again.');
        setIsUploading(false);
      };
      reader.readAsDataURL(file);
    } catch (error) {
      alert('Error uploading logo. Please try again.');
      setIsUploading(false);
    }

    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleRemoveLogo = () => {
    setFormData(prev => ({
      ...prev,
      logoFileName: undefined,
      logoFileData: undefined,
      logoFileType: undefined,
      logoUploadDate: undefined
    }));
  };

  const handleDownloadLogo = () => {
    if (!formData.logoFileData || !formData.logoFileName) return;

    try {
      const link = document.createElement('a');
      link.href = formData.logoFileData;
      link.download = formData.logoFileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading logo. Please try again.');
    }
  };

  const handleSave = () => {
    onSave(formData);
    onClose();
  };

  const handleReset = () => {
    if (window.confirm('Are you sure you want to reset all brand settings to default?')) {
      setFormData({
        companyName: 'Housemaid Management',
        tagline: 'Professional Database System',
        primaryColor: '#2563eb',
        secondaryColor: '#7c3aed',
        copyrightText: '© 2024 Housemaid Management. All rights reserved.'
      });
    }
  };

  const formatFileSize = (base64String: string): string => {
    try {
      const sizeInBytes = (base64String.length * 3) / 4;
      if (sizeInBytes < 1024) return `${sizeInBytes.toFixed(0)} B`;
      if (sizeInBytes < 1024 * 1024) return `${(sizeInBytes / 1024).toFixed(1)} KB`;
      return `${(sizeInBytes / (1024 * 1024)).toFixed(1)} MB`;
    } catch {
      return 'Unknown size';
    }
  };

  const formatDate = (dateString: string): string => {
    return new Date(dateString).toLocaleDateString();
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-2xl max-w-4xl w-full max-h-[90vh] overflow-hidden shadow-2xl">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-purple-600 px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-white bg-opacity-20 rounded-lg">
                <Building className="h-6 w-6 text-white" />
              </div>
              <div>
                <h3 className="text-xl font-semibold text-white">Brand Settings</h3>
                <p className="text-blue-100 text-sm">Customize your application branding</p>
              </div>
            </div>
            <button
              onClick={onClose}
              className="p-2 hover:bg-white hover:bg-opacity-20 rounded-lg transition-colors"
            >
              <X className="h-5 w-5 text-white" />
            </button>
          </div>
        </div>

        <div className="flex">
          {/* Settings Form */}
          <div className="flex-1 p-6 overflow-y-auto max-h-[calc(90vh-100px)]">
            <div className="space-y-8">
              {/* Logo Upload Section */}
              <div>
                <h4 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                  <Image className="h-5 w-5 mr-2 text-blue-600" />
                  Company Logo
                </h4>
                
                {formData.logoFileData ? (
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-6">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center space-x-4">
                        <div className="w-16 h-16 bg-gray-100 rounded-lg overflow-hidden flex items-center justify-center">
                          <img
                            src={formData.logoFileData}
                            alt="Logo preview"
                            className="max-w-full max-h-full object-contain"
                          />
                        </div>
                        <div>
                          <p className="font-medium text-gray-900">{formData.logoFileName}</p>
                          <p className="text-sm text-gray-500">
                            {formatFileSize(formData.logoFileData)} • Uploaded {formData.logoUploadDate ? formatDate(formData.logoUploadDate) : 'Unknown date'}
                          </p>
                        </div>
                      </div>
                      <div className="flex items-center space-x-2">
                        <button
                          onClick={() => setShowPreview(true)}
                          className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          title="Preview Logo"
                        >
                          <Eye className="h-4 w-4" />
                        </button>
                        <button
                          onClick={handleDownloadLogo}
                          className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
                          title="Download Logo"
                        >
                          <Download className="h-4 w-4" />
                        </button>
                        <button
                          onClick={handleRemoveLogo}
                          className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                          title="Remove Logo"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 hover:border-blue-400 transition-colors">
                    <div className="text-center">
                      <div className="mx-auto h-16 w-16 text-gray-400 mb-4">
                        {isUploading ? (
                          <div className="animate-spin rounded-full h-16 w-16 border-b-2 border-blue-600"></div>
                        ) : (
                          <Image className="h-16 w-16" />
                        )}
                      </div>
                      <div className="mb-4">
                        <p className="text-lg font-medium text-gray-900 mb-1">
                          {isUploading ? 'Uploading Logo...' : 'Upload Company Logo'}
                        </p>
                        <p className="text-sm text-gray-500">
                          PNG, JPG, WebP, or SVG up to 1MB
                        </p>
                      </div>
                      <input
                        ref={fileInputRef}
                        type="file"
                        accept=".jpg,.jpeg,.png,.webp,.svg"
                        onChange={handleLogoUpload}
                        disabled={isUploading}
                        className="hidden"
                      />
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        disabled={isUploading}
                        className="bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2 mx-auto"
                      >
                        <Upload className="h-4 w-4" />
                        <span>{isUploading ? 'Uploading...' : 'Choose Logo'}</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Company Information */}
              <div>
                <h4 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                  <Type className="h-5 w-5 mr-2 text-blue-600" />
                  Company Information
                </h4>
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Company Name
                    </label>
                    <input
                      type="text"
                      value={formData.companyName || ''}
                      onChange={(e) => setFormData(prev => ({ ...prev, companyName: e.target.value }))}
                      className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      placeholder="Enter company name"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Tagline
                    </label>
                    <input
                      type="text"
                      value={formData.tagline || ''}
                      onChange={(e) => setFormData(prev => ({ ...prev, tagline: e.target.value }))}
                      className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      placeholder="Enter company tagline"
                    />
                  </div>
                </div>
              </div>

              {/* Copyright Text Section */}
              <div>
                <h4 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                  <Copyright className="h-5 w-5 mr-2 text-blue-600" />
                  Login Page Footer
                </h4>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Copyright Text
                  </label>
                  <input
                    type="text"
                    value={formData.copyrightText || ''}
                    onChange={(e) => setFormData(prev => ({ ...prev, copyrightText: e.target.value }))}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="© 2024 Housemaid Management. All rights reserved."
                  />
                  <p className="text-xs text-gray-500 mt-2">
                    This text will appear at the bottom of the login page. Leave empty to use the default format.
                  </p>
                </div>
              </div>

              {/* Color Scheme */}
              <div>
                <h4 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                  <Palette className="h-5 w-5 mr-2 text-blue-600" />
                  Color Scheme
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Primary Color
                    </label>
                    <div className="flex items-center space-x-3">
                      <input
                        type="color"
                        value={formData.primaryColor || '#2563eb'}
                        onChange={(e) => setFormData(prev => ({ ...prev, primaryColor: e.target.value }))}
                        className="w-12 h-12 border border-gray-300 rounded-lg cursor-pointer"
                      />
                      <input
                        type="text"
                        value={formData.primaryColor || '#2563eb'}
                        onChange={(e) => setFormData(prev => ({ ...prev, primaryColor: e.target.value }))}
                        className="flex-1 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-mono text-sm"
                        placeholder="#2563eb"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Secondary Color
                    </label>
                    <div className="flex items-center space-x-3">
                      <input
                        type="color"
                        value={formData.secondaryColor || '#7c3aed'}
                        onChange={(e) => setFormData(prev => ({ ...prev, secondaryColor: e.target.value }))}
                        className="w-12 h-12 border border-gray-300 rounded-lg cursor-pointer"
                      />
                      <input
                        type="text"
                        value={formData.secondaryColor || '#7c3aed'}
                        onChange={(e) => setFormData(prev => ({ ...prev, secondaryColor: e.target.value }))}
                        className="flex-1 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-mono text-sm"
                        placeholder="#7c3aed"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* Action Buttons */}
              <div className="flex justify-between items-center pt-6 border-t">
                <button
                  onClick={handleReset}
                  className="px-4 py-2 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium flex items-center space-x-2"
                >
                  <RotateCcw className="h-4 w-4" />
                  <span>Reset to Default</span>
                </button>
                <div className="flex space-x-3">
                  <button
                    onClick={onClose}
                    className="px-6 py-3 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleSave}
                    className="px-6 py-3 bg-gradient-to-r from-blue-600 to-purple-600 text-white rounded-lg hover:from-blue-700 hover:to-purple-700 transition-all duration-200 font-medium flex items-center space-x-2"
                  >
                    <Save className="h-4 w-4" />
                    <span>Save Settings</span>
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Live Preview */}
          <div className="w-80 bg-gray-50 border-l p-6">
            <h4 className="text-lg font-semibold text-gray-900 mb-4">Live Preview</h4>
            <div className="space-y-6">
              {/* Header Preview */}
              <div className="bg-white rounded-lg shadow-sm border p-4">
                <h5 className="text-sm font-medium text-gray-700 mb-3">Navigation Header</h5>
                <BrandLogo brandSettings={formData} size="medium" />
              </div>

              {/* Login Page Preview */}
              <div className="bg-white rounded-lg shadow-sm border p-4">
                <h5 className="text-sm font-medium text-gray-700 mb-3">Login Page</h5>
                <div 
                  className="rounded-lg p-4 text-center mb-3"
                  style={{
                    background: `linear-gradient(to right, ${formData.primaryColor || '#2563eb'}, ${formData.secondaryColor || '#7c3aed'})`
                  }}
                >
                  <BrandLogo brandSettings={formData} size="large" />
                </div>
                {/* Copyright Preview */}
                <div className="text-center">
                  <p className="text-xs text-gray-500">
                    {formData.copyrightText || `© 2024 ${formData.companyName || 'Housemaid Management System'}. All rights reserved.`}
                  </p>
                </div>
              </div>

              {/* Small Logo Preview */}
              <div className="bg-white rounded-lg shadow-sm border p-4">
                <h5 className="text-sm font-medium text-gray-700 mb-3">Compact View</h5>
                <BrandLogo brandSettings={formData} size="small" />
              </div>

              {/* Color Palette */}
              <div className="bg-white rounded-lg shadow-sm border p-4">
                <h5 className="text-sm font-medium text-gray-700 mb-3">Color Palette</h5>
                <div className="space-y-2">
                  <div className="flex items-center space-x-3">
                    <div 
                      className="w-6 h-6 rounded border"
                      style={{ backgroundColor: formData.primaryColor || '#2563eb' }}
                    ></div>
                    <span className="text-sm text-gray-600">Primary</span>
                  </div>
                  <div className="flex items-center space-x-3">
                    <div 
                      className="w-6 h-6 rounded border"
                      style={{ backgroundColor: formData.secondaryColor || '#7c3aed' }}
                    ></div>
                    <span className="text-sm text-gray-600">Secondary</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Logo Preview Modal */}
        {showPreview && formData.logoFileData && (
          <div className="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center p-4 z-60">
            <div className="bg-white rounded-lg max-w-2xl w-full max-h-[80vh] overflow-hidden">
              <div className="flex items-center justify-between p-4 border-b">
                <h4 className="text-lg font-semibold text-gray-900">Logo Preview</h4>
                <button
                  onClick={() => setShowPreview(false)}
                  className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
                >
                  <X className="h-5 w-5 text-gray-500" />
                </button>
              </div>
              <div className="p-6 flex justify-center">
                <img
                  src={formData.logoFileData}
                  alt="Logo preview"
                  className="max-w-full max-h-[60vh] object-contain"
                />
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default BrandSettings;