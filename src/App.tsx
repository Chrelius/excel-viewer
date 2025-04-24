import React from 'react';
import ExcelViewer from './components/ExcelViewer'; // Changed to default import

const App: React.FC = () => {
  return (
    <div className="App">
      <h1>Excel Viewer</h1>
      <ExcelViewer />
    </div>
  );
};

export default App;