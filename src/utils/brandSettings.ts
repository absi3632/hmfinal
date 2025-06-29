import { BrandSettings } from '../types/brand';

const BRAND_STORAGE_KEY = 'housemaid_brand_settings';

const DEFAULT_BRAND_SETTINGS: BrandSettings = {
  companyName: 'Housemaid Management',
  tagline: 'Professional Database System',
  primaryColor: '#2563eb', // blue-600
  secondaryColor: '#7c3aed', // purple-600
  copyrightText: '© 2024 Housemaid Management. All rights reserved.',
};

export const saveBrandSettings = (settings: BrandSettings): void => {
  localStorage.setItem(BRAND_STORAGE_KEY, JSON.stringify(settings));
};

export const loadBrandSettings = (): BrandSettings => {
  const data = localStorage.getItem(BRAND_STORAGE_KEY);
  return data ? { ...DEFAULT_BRAND_SETTINGS, ...JSON.parse(data) } : DEFAULT_BRAND_SETTINGS;
};

export const resetBrandSettings = (): void => {
  localStorage.removeItem(BRAND_STORAGE_KEY);
};