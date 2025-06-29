import { User, LoginCredentials } from '../types/user';

const USERS_STORAGE_KEY = 'housemaid_users';
const AUTH_STORAGE_KEY = 'housemaid_auth';

// Default admin user
const DEFAULT_ADMIN: User = {
  id: 'admin-001',
  email: 'stgs.as2@gmail.com',
  password: 'JMasters1!', // In production, this would be hashed
  firstName: 'System',
  lastName: 'Administrator',
  role: 'admin',
  isActive: true,
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
  profileImage: undefined
};

export const initializeUsers = (): void => {
  const existingUsers = localStorage.getItem(USERS_STORAGE_KEY);
  if (!existingUsers) {
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify([DEFAULT_ADMIN]));
  } else {
    const users: User[] = JSON.parse(existingUsers);
    const adminExists = users.some(user => user.email === DEFAULT_ADMIN.email);
    if (!adminExists) {
      users.push(DEFAULT_ADMIN);
      localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users));
    }
  }
};

export const authenticateUser = (credentials: LoginCredentials): User | null => {
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  const user = users.find(u => 
    u.email === credentials.email && 
    u.password === credentials.password && 
    u.isActive
  );
  
  if (user) {
    // Update last login
    user.lastLogin = new Date().toISOString();
    const updatedUsers = users.map(u => u.id === user.id ? user : u);
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(updatedUsers));
    
    // Store auth session
    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
      userId: user.id,
      email: user.email,
      role: user.role,
      loginTime: new Date().toISOString(),
      authMethod: 'password'
    }));
    
    return user;
  }
  
  return null;
};

export const authenticateGoogleUser = (email: string, googleUserData: {
  name: string;
  picture: string;
  given_name: string;
  family_name: string;
}): User | null => {
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  const user = users.find(u => 
    u.email.toLowerCase() === email.toLowerCase() && 
    u.isActive
  );
  
  if (user) {
    // Update last login and Google profile info
    const updatedUser = {
      ...user,
      lastLogin: new Date().toISOString(),
      profileImage: googleUserData.picture, // Update profile image from Google
      // Optionally update name if not set or different
      firstName: user.firstName || googleUserData.given_name,
      lastName: user.lastName || googleUserData.family_name,
    };
    
    const updatedUsers = users.map(u => u.id === user.id ? updatedUser : u);
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(updatedUsers));
    
    // Store auth session
    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
      userId: user.id,
      email: user.email,
      role: user.role,
      loginTime: new Date().toISOString(),
      authMethod: 'google'
    }));
    
    return updatedUser;
  }
  
  return null;
};

export const isEmailInDatabase = (email: string): boolean => {
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  return users.some(u => u.email.toLowerCase() === email.toLowerCase() && u.isActive);
};

export const getCurrentUser = (): User | null => {
  const authData = localStorage.getItem(AUTH_STORAGE_KEY);
  if (!authData) return null;
  
  const { userId } = JSON.parse(authData);
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  return users.find(u => u.id === userId && u.isActive) || null;
};

export const logout = (): void => {
  localStorage.removeItem(AUTH_STORAGE_KEY);
  
  // Sign out from Google if they used Google auth
  if (typeof window !== 'undefined' && window.google) {
    try {
      window.google.accounts.id.disableAutoSelect();
    } catch (error) {
      console.log('Google sign out not available');
    }
  }
};

export const getAllUsers = (): User[] => {
  return JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
};

export const saveUser = (user: User): void => {
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  const existingIndex = users.findIndex(u => u.id === user.id);
  
  if (existingIndex >= 0) {
    users[existingIndex] = { ...user, updatedAt: new Date().toISOString() };
  } else {
    users.push({ ...user, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() });
  }
  
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users));
};

export const deleteUser = (userId: string): void => {
  const users: User[] = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
  const filteredUsers = users.filter(u => u.id !== userId);
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(filteredUsers));
};

export const hasPermission = (userRole: string, requiredRole: string): boolean => {
  const roleHierarchy = { admin: 3, manager: 2, viewer: 1 };
  return roleHierarchy[userRole as keyof typeof roleHierarchy] >= roleHierarchy[requiredRole as keyof typeof roleHierarchy];
};