export const appConfig = {
    version: '1.0.0',
    buildDate: '2025-04-24',
    buildTime: '02:05:17',
    username: process.env.REACT_APP_USERNAME || 'Chrelius',
    isProduction: process.env.NODE_ENV === 'production',
    baseUrl: process.env.NODE_ENV === 'production' ? '/excel-viewer' : ''
  };