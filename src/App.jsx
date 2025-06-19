import React, { useState, useEffect, useRef } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import msdLogo from './assets/msd_logo.webp';

const App = () => {
  const { instance, accounts } = useMsal();
  const [view, setView] = useState('signin');
  const [entity, setEntity] = useState('');
  const [month, setMonth] = useState('');
  const [year, setYear] = useState('');
  const [invoiceData, setInvoiceData] = useState([]);

  const entityOptions = [1207, 3188, 1012, 1194, 380, 519, 1209, 1310, 3124, 1180, 1467, 466, 3121, 477, 1456, 1287,
    1396, 3168, 417, 3583, 1698, 1443, 1662, 1204, 478, 1029,
    1471, 1177, 1253, 1580, 3592, 1285, 3225, 1101, 1395, 1203,
    1247, 1083, 1216, 1190, 3325, 3143, 3223, 1619];
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

    if (!response.ok) alert(`❌ Upload failed: ${response.statusText}`);
  };

  const getDownloadLink = async (fileName) => {
    const token = await getAccessToken();
    const requestUrl = `https://graph.microsoft.com/v1.0/sites/collaboration.merck.com,7c55f2f5-011e-404b-8ab4-2e63558acce8,453db3a9-a975-4499-8e4b-2b358f883ed4/drive/root:/General/PWC Revenue Testing Automation/${fileName}`;

    const response = await fetch(requestUrl, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    if (response.ok) {
      const data = await response.json();
      return data['@microsoft.graph.downloadUrl'];
    } else {
      alert('❌ Failed to fetch download link.');
      return null;
    }
  };

  const handleFileUpload = (e, idx, key) => {
    const file = e.target.files[0];
    if (file) {
      uploadFile(file);
      const updated = [...invoiceData];
      updated[idx][`${key}_file`] = file.name;
      setInvoiceData(updated);
    }
  };

  const downloadFile = async (fileName) => {
    const downloadUrl = await getDownloadLink(fileName);
    if (downloadUrl) window.open(downloadUrl, '_blank');
  };

  const addRow = () => {
    setInvoiceData([...invoiceData, { invoice: '', cash_app: '', credit_note: '', fbl5n: '', cmm: '', comments: '' }]);
  };

  const handlePaste = (e) => {
    const pasteData = e.clipboardData.getData('text');
    const rows = pasteData.split('\n').map(row => row.split('\t'));
    const updatedData = [...invoiceData];

    rows.forEach((cells, rowIndex) => {
      if (!updatedData[rowIndex]) updatedData[rowIndex] = { invoice: '', cash_app: '', credit_note: '', fbl5n: '', cmm: '', comments: '' };
      Object.keys(updatedData[rowIndex]).forEach((key, cellIndex) => {
        if (cells[cellIndex] !== undefined) updatedData[rowIndex][key] = cells[cellIndex];
      });
    });

    setInvoiceData(updatedData);
    e.preventDefault();
  };

  const FileInputCell = ({ value, onTextChange, onFileUpload, fileName }) => {
    const fileRef = useRef();
    return (
      <div style={{ position: 'relative' }}>
        <input
          type="text"
          value={value || ''}
          onChange={onTextChange}
          onClick={() => fileRef.current.click()}
          style={{ width: '100%', padding: '4px', textAlign: 'center' }}
        />
        <input type="file" ref={fileRef} onChange={onFileUpload} style={{ display: 'none' }} />
        {fileName && (
          <span
            onClick={(e) => {
              e.stopPropagation();
              downloadFile(fileName);
            }}
            style={{ cursor: 'pointer', marginLeft: 5, fontSize: 16 }}
            title="Download"
          >
            📥
          </span>
        )}
      </div>
    );
  };

  const headers = ['Invoice', 'Cash App', 'Credit Note', 'FBL5N', 'CMM', 'Comments'];

  return (
    <div style={{ backgroundColor: '#EAF6FC', minHeight: '100vh', padding: '2rem', fontFamily: 'Segoe UI' }}>
      {view !== 'signin' && (
        <img src={msdLogo} alt='MSD Logo' style={{ position: 'absolute', top: 10, right: 20, height: 40 }} />
      )}

      {view === 'signin' && (
        <div style={{ textAlign: 'center', marginTop: 100 }}>
          <img src={msdLogo} alt='MSD Logo' style={{ height: 80, marginBottom: 20 }} />
          <h1>PWC Testing Automation</h1>
          <button onClick={signIn}>Sign in with Microsoft</button>
        </div>
      )}

      {view === 'home' && (
        <div style={{ textAlign: 'center', marginTop: 50 }}>
          {['cash_app', 'po_pod', 'follow_up'].map(s => (
            <button key={s} onClick={() => setView('dashboard')} style={{ margin: 5 }}>
              {s.replace('_', ' ').toUpperCase()}
            </button>
          ))}
          <br />
          <button onClick={logout}>Logout</button>
        </div>
      )}

      {view === 'dashboard' && (
        <div style={{ textAlign: 'center', marginTop: 50 }}>
          <select value={entity} onChange={(e) => setEntity(e.target.value)} style={{ margin: 5 }}>
            <option>-- Entity --</option>
            {entityOptions.map(v => <option key={v}>{v}</option>)}
          </select>
          <select value={month} onChange={(e) => setMonth(e.target.value)} style={{ margin: 5 }}>
            <option>-- Month --</option>
            {months.map(m => <option key={m}>{m}</option>)}
          </select>
          <select value={year} onChange={(e) => setYear(e.target.value)} style={{ margin: 5 }}>
            <option>-- Year --</option>
            {years.map(y => <option key={y}>{y}</option>)}
          </select>
          <br />
          <button onClick={() => setView('upload')} style={{ margin: 10 }}>Submit</button>
        </div>
      )}

      {view === 'upload' && (
        <div style={{ marginTop: 50, overflowX: 'auto' }}>
          <table style={{ width: '100%' }}>
            <thead>
              <tr>{headers.map(h => <th key={h}>{h}</th>)}</tr>
            </thead>
            <tbody onPaste={handlePaste}>
              {invoiceData.map((row, idx) => (
                <tr key={idx}>
                  {headers.map((h, i) => (
                    <td key={i}>
                      <FileInputCell
                        value={row[h.toLowerCase().replace(' ', '_')]}
                        onTextChange={(e) => {
                          const updated = [...invoiceData];
                          updated[idx][h.toLowerCase().replace(' ', '_')] = e.target.value;
                          setInvoiceData(updated);
                        }}
                        onFileUpload={(e) => handleFileUpload(e, idx, h.toLowerCase().replace(' ', '_'))}
                        fileName={row[`${h.toLowerCase().replace(' ', '_')}_file`]}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <button onClick={addRow}>➕ Add Row</button>
        </div>
      )}
    </div>
  );
};

export default App;
