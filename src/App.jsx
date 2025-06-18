import React, { useState, useEffect, useRef } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import msdLogo from './assets/msd_logo.webp';

const App = () => {
  const { instance, accounts } = useMsal();
  const [view, setView] = useState('signin');
  const [section, setSection] = useState('');
  const [entity, setEntity] = useState('');
  const [month, setMonth] = useState('');
  const [year, setYear] = useState('');
  const [invoiceData, setInvoiceData] = useState([]);

  const entityOptions = [1207, 3188, 1012, 1194, 380, 519, 1209, 1310, 3124, 1180, 1467, 466, 3121, 477, 1456, 1287];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const years = ['2025', '2026'];

  useEffect(() => {
    if (accounts.length > 0) setView('home');
  }, [accounts]);

  const signIn = () => instance.loginRedirect(loginRequest);
  const logout = () => instance.logoutRedirect();

  const getAccessToken = async () => {
    const response = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
    return response.accessToken;
  };

  const uploadFile = async (file) => {
    const token = await getAccessToken();
    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/collaboration.merck.com,7c55f2f5-011e-404b-8ab4-2e63558acce8,453db3a9-a975-4499-8e4b-2b358f883ed4/drive/root:/General/PWC Revenue Testing Automation/${file.name}:/content`;

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': file.type,
      },
      body: file
    });

    if (!response.ok) alert(`❌ Upload failed: ${response.statusText}`);
  };

  const downloadFile = async (fileName) => {
    const token = await getAccessToken();
    const downloadUrl = `https://graph.microsoft.com/v1.0/sites/collaboration.merck.com,7c55f2f5-011e-404b-8ab4-2e63558acce8,453db3a9-a975-4499-8e4b-2b358f883ed4/drive/root:/General/PWC Revenue Testing Automation/${fileName}:/content`;

    const response = await fetch(downloadUrl, {
      headers: { 'Authorization': `Bearer ${token}` },
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      link.click();
    } else {
      alert(`❌ Download failed: ${response.statusText}`);
    }
  };

  const handleFileUpload = (e, rowIdx, key) => {
    const file = e.target.files[0];
    if (file) uploadFile(file);
  };

  const addRow = () => {
    setInvoiceData([...invoiceData, { invoice: '', cash_app: '', credit_note: '', fbl5n: '', cmm: '', comments: '' }]);
  };

  const FileInputCell = ({ value, onTextChange, onFileUpload, fileName }) => {
    const fileRef = useRef();
    return (
      <div style={{ position: 'relative', cursor: 'pointer' }}>
        <input
          type="text"
          value={value || ''}
          onChange={onTextChange}
          onClick={() => fileRef.current?.click()}
          style={{ width: '100%', padding: '4px', textAlign: 'center' }}
        />
        <input type="file" ref={fileRef} onChange={onFileUpload} style={{ display: 'none' }} />
        {fileName && (
          <button style={{ fontSize: '0.75rem' }} onClick={() => downloadFile(fileName)}>
            📥 Download
          </button>
        )}
      </div>
    );
  };

  const headers = [
    { key: 'invoice', label: 'Invoice' },
    { key: 'cash_app', label: 'Cash App' },
    { key: 'credit_note', label: 'Credit Note' },
    { key: 'fbl5n', label: 'FBL5N' },
    { key: 'cmm', label: 'CMM' },
    { key: 'comments', label: 'Comments' }
  ];

  return (
    <div style={{ backgroundColor: '#EAF6FC', minHeight: '100vh', fontFamily: 'Segoe UI', padding: '3rem', boxSizing: 'border-box' }}>
      
      {view !== 'signin' && (
        <img src={msdLogo} alt='MSD Logo' style={{ position: 'absolute', top: 15, right: 15, height: 35 }} />
      )}

      {view === 'signin' && (
        <div style={{ textAlign: 'center', marginTop: '5rem' }}>
          <img src={msdLogo} alt='MSD Logo' style={{ height: 80, marginBottom: '1rem' }} />
          <h1>PWC Testing Automation</h1>
          <button style={{ padding: '8px 16px' }} onClick={signIn}>Sign in with Microsoft</button>
        </div>
      )}

      {view === 'home' && (
        <div style={{ textAlign: 'center' }}>
          <h2 style={{ marginBottom: '2rem' }}>Select a section to continue:</h2>
          {['cash_app', 'po_pod', 'follow_up'].map(s => (
            <button key={s} onClick={() => { setSection(s); setView('dashboard'); }} style={{ padding: '12px 25px', margin: '0 10px', backgroundColor: '#007680', color: '#fff', borderRadius: 5 }}>
              {s.replace('_', ' ').toUpperCase()}
            </button>
          ))}
          <button onClick={logout}>Logout</button>
        </div>
      )}

      {view === 'dashboard' && (
        <div style={{ textAlign: 'center', marginTop: '2rem' }}>
          <select value={entity} onChange={(e) => setEntity(e.target.value)} className="dropdown-style">
            <option>-- Select Entity --</option>
            {entityOptions.map((v) => <option key={v}>{v}</option>)}
          </select>
          <select value={month} onChange={(e) => setMonth(e.target.value)} className="dropdown-style">
            <option>-- Select Month --</option>
            {months.map(m => <option key={m}>{m}</option>)}
          </select>
          <select value={year} onChange={(e) => setYear(e.target.value)} className="dropdown-style">
            <option>-- Select Year --</option>
            {years.map(y => <option key={y}>{y}</option>)}
          </select>
          <button onClick={() => setView('upload')}>Submit</button>
          <button onClick={() => setView('home')}>← Go Back</button>
        </div>
      )}

      {view === 'upload' && (
        <div style={{ margin: 'auto', marginTop: '2rem', width: '95%' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '2rem' }}>
            <thead style={{ background: '#007680', color: '#fff' }}>
              <tr>{headers.map(h => <th key={h.key} style={{ padding: 8 }}>{h.label}</th>)}</tr>
            </thead>
            <tbody>
              {invoiceData.map((row, idx) => (
                <tr key={idx}>
                  {headers.map(h => (
                    <td key={h.key}>
                      <FileInputCell
                        value={row[h.key]}
                        onTextChange={(e) => {
                          const updated = [...invoiceData];
                          updated[idx][h.key] = e.target.value;
                          setInvoiceData(updated);
                        }}
                        onFileUpload={(e) => handleFileUpload(e, idx, h.key)}
                        fileName={row[h.key]}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <button onClick={addRow}>➕ Add Row</button>
          <button onClick={() => setView('dashboard')}>← Go Back</button>
        </div>
      )}
    </div>
  );
};

export default App;
