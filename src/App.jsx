import React, { useState, useEffect } from 'react';
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

  const entityOptions = [1207, 3188, 1012, 1194, 380, 519, 1209, 1310, 3124, 1180, 1467, 466, 3121, 477, 1456, 1287,
    1396, 3168, 417, 3583, 1698, 1443, 1662, 1204, 478, 1029,
    1471, 1177, 1253, 1580, 3592, 1285, 3225, 1101, 1395, 1203,
    1247, 1083, 1216, 1190, 3325, 3143, 3223, 1619]; // shortened for brevity
  const months = ['January', 'February', 'March', "April", "May", "June",
  "July", "August", "September", "October", "November", "December"];
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

    if (!response.ok) {
      alert(`❌ Upload failed: ${response.statusText}`);
    }
  };

  const handleFileUpload = (e, rowIdx, key) => {
    const file = e.target.files[0];
    if (file) uploadFile(file);
  };

  const addRow = () => {
    setInvoiceData([...invoiceData, { invoice: '', cash_app: '', credit_note: '', fbl5n: '', cmm: '', comments: '' }]);
  };

  return (
    <div style={{ backgroundColor: '#EAF6FC', minHeight: '100vh', padding: '2rem', fontFamily: 'Segoe UI' }}>
      {view !== 'signin' && (
        <img src={msdLogo} alt='MSD Logo' style={{ position: 'absolute', top: 10, right: 20, height: 40 }} />
      )}

      {view === 'signin' && (
        <div style={{ textAlign: 'center' }}>
          <img src={msdLogo} alt='MSD Logo' style={{ height: 100, marginBottom: 20 }} />
          <h1>PWC Testing Automation</h1>
          <button onClick={signIn}>Sign in with Microsoft</button>
        </div>
      )}

      {view === 'home' && (
        <div>
          <h2>Select a section to continue:</h2>
          {['cash_app', 'po_pod', 'follow_up'].map(s => (
            <button key={s} onClick={() => { setSection(s); setView('dashboard'); }}>
              {s.replace('_', ' ').toUpperCase()}
            </button>
          ))}
          <button onClick={logout}>Logout</button>
        </div>
      )}

      {view === 'dashboard' && (
        <div>
          <select value={entity} onChange={(e) => setEntity(e.target.value)}>
            <option>-- Select Entity --</option>
            {entityOptions.map((v) => <option key={v}>{v}</option>)}
          </select>
          <select value={month} onChange={(e) => setMonth(e.target.value)}>
            <option>-- Select Month --</option>
            {months.map(m => <option key={m}>{m}</option>)}
          </select>
          <select value={year} onChange={(e) => setYear(e.target.value)}>
            <option>-- Select Year --</option>
            {years.map(y => <option key={y}>{y}</option>)}
          </select>
          <button onClick={() => setView('upload')}>Submit</button>
          <button onClick={() => setView('home')}>← Go Back</button>
        </div>
      )}

      {view === 'upload' && (
        <div>
          <table style={{ width: '100%', tableLayout: 'fixed' }}>
            <thead>
              <tr style={{ backgroundColor: '#007680', color: 'white' }}>
                {['Invoice', 'Cash App', 'Credit Note', 'FBL5N', 'CMM', 'Comments'].map(h => <th key={h}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {invoiceData.map((row, idx) => (
                <tr key={idx}>
                  {Object.keys(row).map((key) => (
                    <td key={key}>
                      <input
                        style={{ width: '90%' }}
                        type='text'
                        value={row[key]}
                        onChange={(e) => {
                          const updated = [...invoiceData];
                          updated[idx][key] = e.target.value;
                          setInvoiceData(updated);
                        }}
                        onClick={() => document.getElementById(`file-upload-${idx}-${key}`).click()}
                      />
                      <input
                        type='file'
                        id={`file-upload-${idx}-${key}`}
                        style={{ display: 'none' }}
                        onChange={(e) => handleFileUpload(e, idx, key)}
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
