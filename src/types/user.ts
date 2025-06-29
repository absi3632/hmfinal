export interface User {
  id: string;
  email: string;
  password: string; // In production, this would be hashed
  firstName: string;
  lastName: string;
  role: 'admin' | 'manager' | 'viewer';
  isActive: boolean;
  createdAt: string;
  updatedAt: string;
  lastLogin?: string;
  profileImage?: string;
}

export interface AuthState {
  isAuthenticated: boolean;
  user: User | null;
  loading: boolean;
}

export interface LoginCredentials {
  email: string;
  password: string;
}