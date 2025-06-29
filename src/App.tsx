import { useState, useEffect } from 'react';
import { Housemaid } from './types/housemaid';
import { AuthState, LoginCredentials } from './types/user';
import { BrandSettings as BrandSettingsType } from './types/brand';
import { saveHousemaids, loadHousemaids } from './utils/localStorage';
import { initializeUsers, authenticateUser, getCurrentUser, logout, hasPermission } from './utils/auth';
import { saveBrandSettings, loadBrandSettings } from './utils/brandSettings';
import HousemaidList from './components/HousemaidList';
import HousemaidForm from './components/HousemaidForm';
import Dashboard from './components/Dashboard';
import LoginPage from './components/LoginPage';
import UserManagement from './components/UserManagement';
import BrandSettings from './components/BrandSettings';
import BrandLogo from './components/BrandLogo';
import { LayoutDashboard, Users, Plus, Settings, LogOut, Shield, Building } from 'lucide-react';

function App() {
  const [housemaids, setHousemaids] = useState<Housemaid[]>([]);
  const [currentView, setCurrentView] = useState<'dashboard' | 'list' | 'form' | 'users' | 'brand'>('dashboard');
  const [editingHousemaid, setEditingHousemaid] = useState<Housemaid | undefined>();
  const [authState, setAuthState] = useState<AuthState>({
    isAuthenticated: false,
    user: null,
    loading: true
  });
  const [brandSettings, setBrandSettings] = useState<BrandSettingsType>(loadBrandSettings());
  const [loginError, setLoginError] = useState<string>('');
  const [loginLoading, setLoginLoading] = useState(false);
  const [showBrandSettings, setShowBrandSettings] = useState(false);

  useEffect(() => {
    // Initialize users and check for existing session
    initializeUsers();
    const currentUser = getCurrentUser();
    
    setAuthState({
      isAuthenticated: !!currentUser,
      user: currentUser,
      loading: false
    });

    if (currentUser) {
      const loadedHousemaids = loadHousemaids();
      setHousemaids(loadedHousemaids);
    }

    // Load brand settings
    const loadedBrandSettings = loadBrandSettings();
    setBrandSettings(loadedBrandSettings);
  }, []);

  useEffect(() => {
    if (authState.isAuthenticated) {
      saveHousemaids(housemaids);
    }
  }, [housemaids, authState.isAuthenticated]);

  const handleLogin = async (credentials: LoginCredentials) => {
    setLoginLoading(true);
    setLoginError('');

    // Simulate network delay for better UX
    await new Promise(resolve => setTimeout(resolve, 1000));

    const user = authenticateUser(credentials);
    
    if (user) {
      setAuthState({
        isAuthenticated: true,
        user,
        loading: false
      });
      
      // Load housemaids data after successful login
      const loadedHousemaids = loadHousemaids();
      setHousemaids(loadedHousemaids);
      setCurrentView('dashboard');
    } else {
      setLoginError('Invalid email or password. Please try again.');
    }
    
    setLoginLoading(false);
  };

  const handleLogout = () => {
    logout();
    setAuthState({
      isAuthenticated: false,
      user: null,
      loading: false
    });
    setHousemaids([]);
    setCurrentView('dashboard');
    setEditingHousemaid(undefined);
  };

  const handleBrandSettingsSave = (settings: BrandSettingsType) => {
    setBrandSettings(settings);
    saveBrandSettings(settings);
  };

  const handleAddNew = () => {
    setEditingHousemaid(undefined);
    setCurrentView('form');
  };

  const handleEdit = (housemaid: Housemaid) => {
    setEditingHousemaid(housemaid);
    setCurrentView('form');
  };

  const handleSave = (housemaid: Housemaid) => {
    if (editingHousemaid) {
      setHousemaids(prev => prev.map(h => h.id === housemaid.id ? housemaid : h));
    } else {
      setHousemaids(prev => [...prev, housemaid]);
    }
    setCurrentView('dashboard');
    setEditingHousemaid(undefined);
  };

  const handleDelete = (id: string) => {
    setHousemaids(prev => prev.filter(h => h.id !== id));
  };

  const handleCancel = () => {
    setCurrentView('dashboard');
    setEditingHousemaid(undefined);
  };

  const handleBulkImport = (importedHousemaids: Housemaid[]) => {
    setHousemaids(prev => [...prev, ...importedHousemaids]);
  };

  const handleViewHousemaid = (housemaid: Housemaid) => {
    setEditingHousemaid(housemaid);
    setCurrentView('form');
  };

  // Show loading screen
  if (authState.loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Loading...</p>
        </div>
      </div>
    );
  }

  // Show login page if not authenticated
  if (!authState.isAuthenticated) {
    return (
      <LoginPage 
        onLogin={handleLogin}
        error={loginError}
        loading={loginLoading}
        brandSettings={brandSettings}
      />
    );
  }

  const navigationItems = [
    {
      id: 'dashboard',
      label: 'Dashboard',
      icon: <LayoutDashboard className="h-5 w-5" />,
      view: 'dashboard' as const,
      permission: 'viewer'
    },
    {
      id: 'records',
      label: 'Records',
      icon: <Users className="h-5 w-5" />,
      view: 'list' as const,
      permission: 'viewer'
    },
    {
      id: 'users',
      label: 'User Management',
      icon: <Settings className="h-5 w-5" />,
      view: 'users' as const,
      permission: 'admin'
    }
  ].filter(item => hasPermission(authState.user!.role, item.permission));

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Navigation Header */}
      <div className="bg-white shadow-sm border-b sticky top-0 z-40">
        <div className="container mx-auto px-4">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center space-x-8">
              <BrandLogo brandSettings={brandSettings} size="medium" />
              
              <nav className="flex space-x-1">
                {navigationItems.map((item) => (
                  <button
                    key={item.id}
                    onClick={() => setCurrentView(item.view)}
                    className={`flex items-center space-x-2 px-4 py-2 rounded-lg font-medium transition-all duration-200 ${
                      currentView === item.view
                        ? 'bg-blue-100 text-blue-700 shadow-sm'
                        : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                    }`}
                  >
                    {item.icon}
                    <span>{item.label}</span>
                  </button>
                ))}
              </nav>
            </div>

            <div className="flex items-center space-x-4">
              {hasPermission(authState.user!.role, 'manager') && (
                <button
                  onClick={handleAddNew}
                  className="bg-gradient-to-r from-blue-600 to-purple-600 text-white px-6 py-2 rounded-lg hover:from-blue-700 hover:to-purple-700 transition-all duration-200 font-medium flex items-center space-x-2 shadow-sm hover:shadow-md"
                  style={{
                    background: `linear-gradient(to right, ${brandSettings.primaryColor || '#2563eb'}, ${brandSettings.secondaryColor || '#7c3aed'})`
                  }}
                >
                  <Plus className="h-4 w-4" />
                  <span>Add Record</span>
                </button>
              )}

              {/* User Menu */}
              <div className="flex items-center space-x-3">
                <div className="flex items-center space-x-2 px-3 py-2 bg-gray-100 rounded-lg">
                  {authState.user!.profileImage ? (
                    <img
                      src={authState.user!.profileImage}
                      alt={`${authState.user!.firstName} ${authState.user!.lastName}`}
                      className="w-8 h-8 rounded-full object-cover"
                    />
                  ) : (
                    <div 
                      className="w-8 h-8 rounded-full flex items-center justify-center text-white text-sm font-semibold"
                      style={{
                        background: `linear-gradient(to right, ${brandSettings.primaryColor || '#2563eb'}, ${brandSettings.secondaryColor || '#7c3aed'})`
                      }}
                    >
                      {authState.user!.firstName.charAt(0)}{authState.user!.lastName.charAt(0)}
                    </div>
                  )}
                  <div className="text-sm">
                    <p className="font-medium text-gray-900">{authState.user!.firstName} {authState.user!.lastName}</p>
                    <div className="flex items-center space-x-1">
                      <Shield className="h-3 w-3 text-gray-500" />
                      <span className="text-gray-500 capitalize">{authState.user!.role}</span>
                    </div>
                  </div>
                </div>
                
                {hasPermission(authState.user!.role, 'admin') && (
                  <button
                    onClick={() => setShowBrandSettings(true)}
                    className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                    title="Brand Settings"
                  >
                    <Building className="h-5 w-5" />
                  </button>
                )}
                
                <button
                  onClick={handleLogout}
                  className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                  title="Logout"
                >
                  <LogOut className="h-5 w-5" />
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="container mx-auto px-4 py-8">
        {currentView === 'dashboard' && (
          <Dashboard
            housemaids={housemaids}
            onViewHousemaid={handleViewHousemaid}
          />
        )}
        
        {currentView === 'list' && (
          <HousemaidList
            housemaids={housemaids}
            onAdd={handleAddNew}
            onEdit={handleEdit}
            onDelete={handleDelete}
            onBulkImport={handleBulkImport}
          />
        )}
        
        {currentView === 'form' && (
          <HousemaidForm
            housemaid={editingHousemaid}
            onSave={handleSave}
            onCancel={handleCancel}
          />
        )}

        {currentView === 'users' && (
          <UserManagement currentUser={authState.user!} />
        )}
      </div>

      {/* Brand Settings Modal */}
      {showBrandSettings && (
        <BrandSettings
          brandSettings={brandSettings}
          onSave={handleBrandSettingsSave}
          onClose={() => setShowBrandSettings(false)}
        />
      )}
    </div>
  );
}

export default App;