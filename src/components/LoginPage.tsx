import React, { useState } from 'react';
import { Eye, EyeOff, Lock, Mail, Shield, AlertCircle, CheckCircle } from 'lucide-react';
import { LoginCredentials } from '../types/user';
import { BrandSettings } from '../types/brand';
import BrandLogo from './BrandLogo';

interface LoginPageProps {
  onLogin: (credentials: LoginCredentials) => void;
  error?: string;
  loading?: boolean;
  brandSettings: BrandSettings;
}

const LoginPage: React.FC<LoginPageProps> = ({ 
  onLogin, 
  error, 
  loading, 
  brandSettings 
}) => {
  const [credentials, setCredentials] = useState<LoginCredentials>({
    email: '',
    password: ''
  });
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onLogin(credentials);
  };

  const handleInputChange = (field: keyof LoginCredentials, value: string) => {
    setCredentials(prev => ({ ...prev, [field]: value }));
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50 flex items-center justify-center p-4">
      {/* Background Pattern */}
      <div className="absolute inset-0 bg-grid-pattern opacity-5"></div>
      
      <div className="relative w-full max-w-md">
        {/* Login Card */}
        <div className="bg-white rounded-2xl shadow-2xl border border-gray-100 overflow-hidden">
          {/* Header */}
          <div 
            className="px-8 py-10 text-center text-white"
            style={{
              background: `linear-gradient(to right, ${brandSettings.primaryColor || '#2563eb'}, ${brandSettings.secondaryColor || '#7c3aed'})`
            }}
          >
            <div className="w-20 h-20 bg-white bg-opacity-20 rounded-2xl flex items-center justify-center mx-auto mb-4 backdrop-blur-sm overflow-hidden">
              <BrandLogo 
                brandSettings={{
                  ...brandSettings,
                  logoFileData: brandSettings.logoFileData ? brandSettings.logoFileData : undefined
                }} 
                size="large" 
                showText={false}
              />
            </div>
            <h1 className="text-2xl font-bold mb-2">
              {brandSettings.companyName || 'Housemaid Management'}
            </h1>
            <p className="text-blue-100 text-sm">
              {brandSettings.tagline || 'Professional Database System'}
            </p>
          </div>

          {/* Login Form */}
          <div className="px-8 py-8">
            <div className="text-center mb-8">
              <h2 className="text-2xl font-bold text-gray-900 mb-2">Welcome Back</h2>
              <p className="text-gray-600">Please sign in to your account</p>
            </div>

            {/* Error Messages */}
            {error && (
              <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg flex items-center space-x-3">
                <AlertCircle className="h-5 w-5 text-red-600 flex-shrink-0" />
                <p className="text-red-700 text-sm">{error}</p>
              </div>
            )}

            <form onSubmit={handleSubmit} className="space-y-6">
              {/* Email Field */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Email Address
                </label>
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Mail className="h-5 w-5 text-gray-400" />
                  </div>
                  <input
                    type="email"
                    value={credentials.email}
                    onChange={(e) => handleInputChange('email', e.target.value)}
                    className="w-full pl-10 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:border-blue-500 transition-colors bg-gray-50 focus:bg-white"
                    style={{
                      '--tw-ring-color': brandSettings.primaryColor || '#2563eb'
                    } as React.CSSProperties}
                    placeholder="Enter your email"
                    required
                    disabled={loading}
                  />
                </div>
              </div>

              {/* Password Field */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Password
                </label>
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Lock className="h-5 w-5 text-gray-400" />
                  </div>
                  <input
                    type={showPassword ? 'text' : 'password'}
                    value={credentials.password}
                    onChange={(e) => handleInputChange('password', e.target.value)}
                    className="w-full pl-10 pr-12 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:border-blue-500 transition-colors bg-gray-50 focus:bg-white"
                    style={{
                      '--tw-ring-color': brandSettings.primaryColor || '#2563eb'
                    } as React.CSSProperties}
                    placeholder="Enter your password"
                    required
                    disabled={loading}
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute inset-y-0 right-0 pr-3 flex items-center"
                    disabled={loading}
                  >
                    {showPassword ? (
                      <EyeOff className="h-5 w-5 text-gray-400 hover:text-gray-600" />
                    ) : (
                      <Eye className="h-5 w-5 text-gray-400 hover:text-gray-600" />
                    )}
                  </button>
                </div>
              </div>

              {/* Remember Me */}
              <div className="flex items-center justify-between">
                <div className="flex items-center">
                  <input
                    id="remember-me"
                    type="checkbox"
                    checked={rememberMe}
                    onChange={(e) => setRememberMe(e.target.checked)}
                    className="h-4 w-4 border-gray-300 rounded"
                    style={{
                      color: brandSettings.primaryColor || '#2563eb'
                    }}
                    disabled={loading}
                  />
                  <label htmlFor="remember-me" className="ml-2 block text-sm text-gray-700">
                    Remember me
                  </label>
                </div>
                <button
                  type="button"
                  className="text-sm font-medium hover:opacity-80"
                  style={{
                    color: brandSettings.primaryColor || '#2563eb'
                  }}
                  disabled={loading}
                >
                  Forgot password?
                </button>
              </div>

              {/* Submit Button */}
              <button
                type="submit"
                disabled={loading || !credentials.email || !credentials.password}
                className="w-full text-white py-3 px-4 rounded-lg font-medium hover:opacity-90 focus:outline-none focus:ring-2 focus:ring-offset-2 transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center space-x-2"
                style={{
                  background: `linear-gradient(to right, ${brandSettings.primaryColor || '#2563eb'}, ${brandSettings.secondaryColor || '#7c3aed'})`,
                  '--tw-ring-color': brandSettings.primaryColor || '#2563eb'
                } as React.CSSProperties}
              >
                {loading ? (
                  <>
                    <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
                    <span>Signing in...</span>
                  </>
                ) : (
                  <>
                    <Shield className="h-5 w-5" />
                    <span>Sign In</span>
                  </>
                )}
              </button>
            </form>

            {/* Features */}
            <div className="mt-8 pt-6 border-t border-gray-200">
              <p className="text-center text-sm text-gray-600 mb-4">System Features</p>
              <div className="grid grid-cols-2 gap-4 text-xs">
                <div className="flex items-center space-x-2 text-gray-600">
                  <CheckCircle className="h-4 w-4 text-green-500" />
                  <span>Role-based Access</span>
                </div>
                <div className="flex items-center space-x-2 text-gray-600">
                  <CheckCircle className="h-4 w-4 text-green-500" />
                  <span>Secure Authentication</span>
                </div>
                <div className="flex items-center space-x-2 text-gray-600">
                  <CheckCircle className="h-4 w-4 text-green-500" />
                  <span>Data Management</span>
                </div>
                <div className="flex items-center space-x-2 text-gray-600">
                  <CheckCircle className="h-4 w-4 text-green-500" />
                  <span>Report Generation</span>
                </div>
              </div>
            </div>

            {/* Security Notice */}
            <div className="mt-6 p-3 bg-blue-50 border border-blue-200 rounded-lg">
              <div className="flex items-start space-x-2">
                <Shield className="h-4 w-4 text-blue-600 mt-0.5 flex-shrink-0" />
                <div className="text-xs text-blue-700">
                  <p className="font-medium mb-1">Secure Access</p>
                  <p>Only authorized users can access this system. Contact your administrator if you need access.</p>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Footer with Custom Copyright Text */}
        <div className="text-center mt-8">
          <p className="text-sm text-gray-500">
            {brandSettings.copyrightText || `© 2024 ${brandSettings.companyName || 'Housemaid Management System'}. All rights reserved.`}
          </p>
        </div>
      </div>

      {/* Background Decorations */}
      <div 
        className="absolute top-10 left-10 w-20 h-20 rounded-full opacity-20 animate-pulse"
        style={{ backgroundColor: brandSettings.primaryColor || '#3b82f6' }}
      ></div>
      <div 
        className="absolute bottom-10 right-10 w-32 h-32 rounded-full opacity-20 animate-pulse delay-1000"
        style={{ backgroundColor: brandSettings.secondaryColor || '#8b5cf6' }}
      ></div>
      <div 
        className="absolute top-1/2 left-5 w-16 h-16 rounded-full opacity-20 animate-pulse delay-500"
        style={{ backgroundColor: brandSettings.primaryColor || '#6366f1' }}
      ></div>
    </div>
  );
};

export default LoginPage;