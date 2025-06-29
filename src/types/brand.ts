export interface BrandSettings {
  logoFileName?: string;
  logoFileData?: string; // Base64 encoded image data
  logoFileType?: string; // MIME type
  logoUploadDate?: string;
  companyName?: string;
  tagline?: string;
  primaryColor?: string;
  secondaryColor?: string;
  copyrightText?: string; // Custom copyright text for login page
}