// Google Authentication utilities
export interface GoogleUser {
  email: string;
  name: string;
  picture: string;
  given_name: string;
  family_name: string;
}

export interface GoogleAuthResponse {
  credential: string;
  select_by: string;
}

// Google Client ID - Updated with your provided client ID
const GOOGLE_CLIENT_ID = "909478597009-dats95albpo2uhsh2v9ufr131fu0socs.apps.googleusercontent.com";

export const initializeGoogleAuth = (): Promise<void> => {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.google) {
      reject(new Error('Google API not loaded'));
      return;
    }

    try {
      window.google.accounts.id.initialize({
        client_id: GOOGLE_CLIENT_ID,
        callback: () => {}, // Will be set by individual components
        auto_select: false,
        cancel_on_tap_outside: true,
      });
      resolve();
    } catch (error) {
      reject(error);
    }
  });
};

export const parseGoogleCredential = (credential: string): GoogleUser | null => {
  try {
    // Decode JWT token (simplified - in production use a proper JWT library)
    const base64Url = credential.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );

    const payload = JSON.parse(jsonPayload);
    
    return {
      email: payload.email,
      name: payload.name,
      picture: payload.picture,
      given_name: payload.given_name,
      family_name: payload.family_name,
    };
  } catch (error) {
    console.error('Error parsing Google credential:', error);
    return null;
  }
};

export const renderGoogleSignInButton = (
  elementId: string,
  callback: (response: GoogleAuthResponse) => void
): void => {
  if (typeof window === 'undefined' || !window.google) {
    console.error('Google API not loaded');
    return;
  }

  // Set the callback for this specific instance
  window.google.accounts.id.initialize({
    client_id: GOOGLE_CLIENT_ID,
    callback: callback,
    auto_select: false,
    cancel_on_tap_outside: true,
  });

  window.google.accounts.id.renderButton(
    document.getElementById(elementId),
    {
      theme: 'outline',
      size: 'large',
      width: '100%',
      text: 'signin_with',
      shape: 'rectangular',
      logo_alignment: 'left',
    }
  );
};

// Type declarations for Google API
declare global {
  interface Window {
    google: {
      accounts: {
        id: {
          initialize: (config: any) => void;
          renderButton: (element: HTMLElement | null, config: any) => void;
          prompt: () => void;
          disableAutoSelect: () => void;
        };
      };
    };
  }
}