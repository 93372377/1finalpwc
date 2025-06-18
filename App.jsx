import React, { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';

const App = () => {
  const { instance, accounts } = useMsal();
  const [view, setView] = useState('signin');
  const [section, setSection] = useState('');
  const [entity, setEntity] = useState('');
  const [month, setMonth] = useState('');
  const [year, setYear] = useState('');
  const [invoiceData, setInvoiceData] = useState([{}]);
  const [poPodData, setPoPodData] = useState([{}]);
  const [followUpData, setFollowUpData] = useState([{}]);

  const entityOptions = [1207, 3188, 1012, 1194, 380, 519, 1209, 1310, 3124, 1180, 1467, 466, 3121, 477, 1456, 1287, 1396, 3168, 417, 3583, 1698, 1443, 1662, 1204, 478, 1029, 1471, 1177, 1253, 1580, 3592, 1285, 3225, 1101, 1395, 1203, 1247, 1083, 1216, 1190, 3325, 3143, 3223, 1619];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const years = ['2025', '2026'];

  useEffect(() => { if (accounts.length > 0) setView('home'); }, [accounts]);

  const signIn = () => instance.loginRedirect(loginRequest);
  const logout = () => instance.logoutRedirect();

  const getAccessToken = async () => {
    const account = accounts[0];
    const response = await instance.acquireTokenSilent({ ...loginRequest, account });
    return response.accessToken;
  };

  const handleFileUpload = async (e, rowIdx, key, data, setData) => {
    const file = e.target.files[0];
    if (!file) return;

    const accessToken = await getAccessToken();
    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/collaboration.merck.com,7c55f2f5-011e-404b-8ab4-2e63558acce8,453db3a9-a975-4499-8e4b-2b358f883ed4/drive/root:/General/PWC Revenue Testing Automation/${encodeURIComponent(file.name)}:/content`;

    try {
      const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': file.type },
        body: file
      });

      if (response.ok) {
        const updated = [...data];
        updated[rowIdx] = { ...updated[rowIdx], [key]: file.name };
        setData(updated);
      } else {
        alert(`Upload failed: ${response.status}`);
      }
    } catch (err) {
      alert(`Upload request failed: ${err.message}`);
    }
  };

  const renderTable = (headers, data, setData) => (
    <div style={{ backgroundColor: '#E8F6FC', padding: '2rem', position: 'relative' }}>
      <img src="https://logowik.com/content/uploads/images/merck-sharp-dohme-msd5762.logowik.com.webp" style={{ position: 'absolute', top: '10px', right: '10px', height: '50px' }} alt="MSD" />
      <table>
        <thead>
          <tr>
            {headers.map(h => (
              <th key={h.key}>
                {h.label}
                <select><option>All</option></select>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row, idx) => (
            <tr key={idx}>
              {headers.map(h => (
                <td key={h.key}>
                  <input
                    value={row[h.key] || ''}
                    onChange={e => {
                      const updated = [...data];
                      updated[idx][h.key] = e.target.value;
                      setData(updated);
                    }}
                    onClick={() => document.getElementById(`file-${idx}-${h.key}`).click()}
                  />
                  <input type='file' hidden id={`file-${idx}-${h.key}`} onChange={e => handleFileUpload(e, idx, h.key, data, setData)} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <button onClick={() => setData([...data, {}])}>＋ Add Row</button>
      <button onClick={() => setView('dashboard')}>← Go Back</button>
    </div>
  );

  const headersMap = {
    cash_app: ['Invoice', 'Cash App', 'Credit Note', 'FBL5N', 'CMM', 'Comments'],
    po_pod: ['SO', 'PO', 'PO Date', 'POD', 'POD Date', 'Invoice Date', 'Order Creator', 'Plant', 'Customer', 'Product', 'Incoterms'],
    follow_up: ['Group/Statutory', 'Country', 'AH/HH', 'Entity', 'Month', 'SO', 'Invoice', 'POD', 'PO', 'Order Creator', 'Plant', 'Customer', 'Product', 'Year', 'PwC Comment']
  };

  const dataMap = { cash_app: [invoiceData, setInvoiceData], po_pod: [poPodData, setPoPodData], follow_up: [followUpData, setFollowUpData] };

  return (
    <div style={{ backgroundColor: '#E8F6FC', padding: '2rem', minHeight: '100vh', fontFamily: 'Segoe UI' }}>
      {view === 'signin' && (
        <div style={{ textAlign: 'center' }}>
          <img src="https://logowik.com/content/uploads/images/merck-sharp-dohme-msd5762.logowik.com.webp" style={{ height: '150px' }} alt="MSD" />
          <h1>PWC Testing Automation</h1>
          <button onClick={signIn}>Sign in with Microsoft</button>
        </div>
      )}
      {view === 'home' && (
        <div>
          <img src="https://logowik.com/content/uploads/images/merck-sharp-dohme-msd5762.logowik.com.webp" style={{ position: 'absolute', top: '10px', right: '10px', height: '50px' }} alt="MSD" />
          <h2>Select a section to continue:</h2>
          {['cash_app', 'po_pod', 'follow_up'].map(s => <button key={s} onClick={() => { setSection(s); setView('dashboard'); }}>{s.replace('_', ' ').toUpperCase()}</button>)}
        </div>
      )}
      {view === 'dashboard' && (
        <div>
          <img src="https://logowik.com/content/uploads/images/merck-sharp-dohme-msd5762.logowik.com.webp" style={{ position: 'absolute', top: '10px', right: '10px', height: '50px' }} alt="MSD" />
          <form onSubmit={(e) => { e.preventDefault(); setView('upload'); }}>
            <select value={entity} onChange={e => setEntity(e.target.value)} required><option>-- Entity --</option>{entityOptions.map(o => <option key={o}>{o}</option>)}</select>
            <select value={month} onChange={e => setMonth(e.target.value)} required><option>-- Month --</option>{months.map(m => <option key={m}>{m}</option>)}</select>
            <select value={year} onChange={e => setYear(e.target.value)} required><option>-- Year --</option>{years.map(y => <option key={y}>{y}</option>)}</select>
            <button type='submit'>Submit</button>
          </form>
          <button onClick={() => setView('home')}>← Go Back</button>
        </div>
      )}
      {view === 'upload' && renderTable(headersMap[section].map(h => ({ key: h.toLowerCase().replace(/ /g, '_'), label: h })), ...dataMap[section])}
    </div>
  );
};

export default App;
