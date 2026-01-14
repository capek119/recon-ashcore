import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx'; // FIX C: XLSX SUPPORT

// ============================================================================
// RECONPREP V1 - Bank Reconciliation Preparation Tool
// ============================================================================

// Storage keys for localStorage
// V2: Removed USERS and CURRENT_USER, added PREPARED_BY
const STORAGE_KEYS = {
  COMPANY: 'reconprep_company',
  PREPARED_BY: 'reconprep_prepared_by',
  BANK_ACCOUNTS: 'reconprep_bank_accounts',
  PERIODS: 'reconprep_periods',
  MATCHING_CONFIG: 'reconprep_matching_config'
};

// Save to localStorage
const saveToStorage = (key, data) => {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    console.warn('Failed to save to localStorage:', e);
  }
};

// Load from localStorage
const loadFromStorage = (key, defaultValue = null) => {
  try {
    const item = localStorage.getItem(key);
    return item ? JSON.parse(item) : defaultValue;
  } catch (e) {
    console.warn('Failed to load from localStorage:', e);
    return defaultValue;
  }
};

// UUID Generator
const generateId = () => crypto.randomUUID ? crypto.randomUUID() : 
  'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });

// Hash function for data integrity
const hashData = async (data) => {
  const encoder = new TextEncoder();
  const dataBuffer = encoder.encode(JSON.stringify(data));
  const hashBuffer = await crypto.subtle.digest('SHA-256', dataBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('').substring(0, 8);
};

// PDF Generation using browser print
// V2: Changed currentUser to preparedByName
const generatePDFContent = (data) => {
  const { company, account, period, bankLines, cashLines, matches, notes, preparedByName, confirmedItems, stats } = data;
  
  const unmatchedBank = bankLines.filter(b => !matches.find(m => m.bankLineId === b.id));
  const unmatchedCash = cashLines.filter(c => !matches.find(m => m.cashLineId === c.id));
  const bankTotal = bankLines.reduce((sum, l) => sum + l.amount, 0);
  const cashTotal = cashLines.reduce((sum, l) => sum + l.amount, 0);
  const unmatchedBankTotal = unmatchedBank.reduce((sum, l) => sum + l.amount, 0);
  const unmatchedCashTotal = unmatchedCash.reduce((sum, l) => sum + l.amount, 0);
  
  const formatMoney = (amt) => new Intl.NumberFormat('en-MY', { style: 'currency', currency: company?.currency || 'MYR', minimumFractionDigits: 2 }).format(amt || 0);
  const formatDate = (d) => d ? new Date(d).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${company?.companyName} - Bank Reconciliation - ${period?.periodLabel}</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; font-size: 11px; color: #333; padding: 20px; }
        .page { page-break-after: always; min-height: 100vh; }
        .page:last-child { page-break-after: avoid; }
        .header { text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #333; }
        .header h1 { font-size: 18px; margin-bottom: 5px; }
        .header h2 { font-size: 14px; font-weight: normal; color: #666; }
        .section { margin-bottom: 25px; }
        .section-title { font-size: 13px; font-weight: bold; background: #f0f0f0; padding: 8px; margin-bottom: 10px; border-left: 4px solid #333; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        th, td { padding: 6px 8px; text-align: left; border: 1px solid #ddd; }
        th { background: #f5f5f5; font-weight: bold; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .mono { font-family: 'Courier New', monospace; }
        .total-row { font-weight: bold; background: #f9f9f9; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
        .info-box { border: 1px solid #ddd; padding: 15px; border-radius: 4px; }
        .info-box label { font-weight: bold; color: #666; font-size: 10px; display: block; margin-bottom: 3px; }
        .info-box span { font-size: 12px; }
        .recon-table { max-width: 500px; margin: 0 auto; }
        .recon-table td { padding: 8px 12px; }
        .recon-table .indent { padding-left: 30px; color: #666; }
        .recon-table .total { border-top: 2px solid #333; font-weight: bold; }
        .footer { position: fixed; bottom: 20px; left: 20px; right: 20px; font-size: 9px; color: #666; border-top: 1px solid #ddd; padding-top: 10px; display: flex; justify-content: space-between; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 9px; }
        .badge-purple { background: #e9d5ff; color: #7c3aed; }
        .positive { color: #059669; }
        .negative { color: #dc2626; }
        @media print {
          body { padding: 0; }
          .footer { position: fixed; }
        }
      </style>
    </head>
    <body>
      <!-- Page 1: Cover & Summary -->
      <div class="page">
        <div class="header">
          <h1>${company?.companyName || 'Company'}</h1>
          <h2>Bank Reconciliation Statement</h2>
          <p style="margin-top: 10px; font-size: 12px;">
            ${account?.bankName} - ${account?.nickname}<br>
            Period: ${period?.periodLabel}
          </p>
        </div>
        
        <div class="info-grid">
          <div class="info-box">
            <label>Prepared By</label>
            <span>${preparedByName || '-'}</span>
          </div>
          <div class="info-box">
            <label>Date Prepared</label>
            <span>${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
          </div>
          <div class="info-box">
            <label>Company Registration</label>
            <span>${company?.companyRegNo || '-'}</span>
          </div>
          <div class="info-box">
            <label>Account Reference</label>
            <span>${account?.accountRef || '-'}</span>
          </div>
        </div>
        
        <div class="section">
          <div class="section-title">Reconciliation Summary</div>
          <table class="recon-table">
            <tr>
              <td><strong>Balance per Bank Statement</strong></td>
              <td class="text-right mono">${formatMoney(bankTotal)}</td>
            </tr>
            <tr>
              <td class="indent">Less: Outstanding Items (Cash Book)</td>
              <td class="text-right mono negative">${formatMoney(unmatchedCashTotal)}</td>
            </tr>
            <tr class="total">
              <td>Adjusted Bank Balance</td>
              <td class="text-right mono">${formatMoney(bankTotal + unmatchedCashTotal)}</td>
            </tr>
            <tr><td colspan="2" style="height: 15px;"></td></tr>
            <tr>
              <td><strong>Balance per Cash Book</strong></td>
              <td class="text-right mono">${formatMoney(cashTotal)}</td>
            </tr>
            <tr>
              <td class="indent">Add/Less: Unrecorded Bank Items</td>
              <td class="text-right mono">${formatMoney(unmatchedBankTotal)}</td>
            </tr>
            <tr class="total">
              <td>Adjusted Cash Book Balance</td>
              <td class="text-right mono">${formatMoney(cashTotal + unmatchedBankTotal)}</td>
            </tr>
          </table>
        </div>
        
        <div class="section">
          <div class="section-title">Matching Statistics</div>
          <table>
            <tr>
              <th>Metric</th>
              <th class="text-right">Count</th>
              <th class="text-right">Amount</th>
            </tr>
            <tr>
              <td>Total Bank Statement Lines</td>
              <td class="text-right">${bankLines.length}</td>
              <td class="text-right mono">${formatMoney(bankTotal)}</td>
            </tr>
            <tr>
              <td>Total Cash Book Lines</td>
              <td class="text-right">${cashLines.length}</td>
              <td class="text-right mono">${formatMoney(cashTotal)}</td>
            </tr>
            <tr>
              <td>Matched Pairs</td>
              <td class="text-right">${matches.length}</td>
              <td class="text-right">-</td>
            </tr>
            <tr>
              <td>Unmatched Bank Items</td>
              <td class="text-right">${unmatchedBank.length}</td>
              <td class="text-right mono">${formatMoney(unmatchedBankTotal)}</td>
            </tr>
            <tr>
              <td>Unmatched Cash Book Items</td>
              <td class="text-right">${unmatchedCash.length}</td>
              <td class="text-right mono">${formatMoney(unmatchedCashTotal)}</td>
            </tr>
            <tr class="total-row">
              <td>Match Rate</td>
              <td class="text-right" colspan="2">${bankLines.length > 0 ? Math.round((matches.length / bankLines.length) * 100) : 0}%</td>
            </tr>
          </table>
        </div>
      </div>
      
      <!-- Page 2: Unmatched Bank Items -->
      <div class="page">
        <div class="section-title">Unmatched Bank Statement Items (${unmatchedBank.length})</div>
        ${unmatchedBank.length === 0 ? '<p style="padding: 20px; text-align: center; color: #666;">No unmatched bank statement items</p>' : `
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Date</th>
              <th>Description</th>
              <th>Reference</th>
              <th class="text-right">Amount</th>
              <th>Category</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody>
            ${unmatchedBank.map((line, idx) => {
              const note = notes.find(n => n.bankLineId === line.id);
              return `
                <tr>
                  <td>${idx + 1}</td>
                  <td>${formatDate(line.txnDate)}</td>
                  <td>${line.description}</td>
                  <td>${line.reference || '-'}</td>
                  <td class="text-right mono ${line.amount >= 0 ? 'positive' : 'negative'}">${formatMoney(line.amount)}</td>
                  <td>${note ? `<span class="badge badge-purple">${note.category.replace(/_/g, ' ')}</span>` : '-'}</td>
                  <td>${note?.text || '-'}</td>
                </tr>
              `;
            }).join('')}
            <tr class="total-row">
              <td colspan="4"><strong>Total Unmatched Bank Items</strong></td>
              <td class="text-right mono">${formatMoney(unmatchedBankTotal)}</td>
              <td colspan="2"></td>
            </tr>
          </tbody>
        </table>
        `}
      </div>
      
      <!-- Page 3: Unmatched Cash Book Items -->
      <div class="page">
        <div class="section-title">Unmatched Cash Book Items (${unmatchedCash.length})</div>
        ${unmatchedCash.length === 0 ? '<p style="padding: 20px; text-align: center; color: #666;">No unmatched cash book items</p>' : `
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Date</th>
              <th>Description</th>
              <th>Reference</th>
              <th class="text-right">Amount</th>
              <th>Category</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody>
            ${unmatchedCash.map((line, idx) => {
              const note = notes.find(n => n.cashLineId === line.id);
              return `
                <tr>
                  <td>${idx + 1}</td>
                  <td>${formatDate(line.entryDate)}</td>
                  <td>${line.description}</td>
                  <td>${line.reference || '-'}</td>
                  <td class="text-right mono ${line.amount >= 0 ? 'positive' : 'negative'}">${formatMoney(line.amount)}</td>
                  <td>${note ? `<span class="badge badge-purple">${note.category.replace(/_/g, ' ')}</span>` : '-'}</td>
                  <td>${note?.text || '-'}</td>
                </tr>
              `;
            }).join('')}
            <tr class="total-row">
              <td colspan="4"><strong>Total Unmatched Cash Book Items</strong></td>
              <td class="text-right mono">${formatMoney(unmatchedCashTotal)}</td>
              <td colspan="2"></td>
            </tr>
          </tbody>
        </table>
        `}
      </div>
      
      <!-- Page 4: Matched Items -->
      <div class="page">
        <div class="section-title">Matched Transactions (${matches.length} pairs)</div>
        ${matches.length === 0 ? '<p style="padding: 20px; text-align: center; color: #666;">No matched transactions</p>' : `
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Bank Date</th>
              <th>Bank Description</th>
              <th class="text-right">Amount</th>
              <th>Cash Date</th>
              <th>Cash Description</th>
              <th>Match Type</th>
            </tr>
          </thead>
          <tbody>
            ${matches.map((match, idx) => {
              const bank = bankLines.find(b => b.id === match.bankLineId);
              const cash = cashLines.find(c => c.id === match.cashLineId);
              return `
                <tr>
                  <td>${idx + 1}</td>
                  <td>${formatDate(bank?.txnDate)}</td>
                  <td>${bank?.description || '-'}</td>
                  <td class="text-right mono">${formatMoney(bank?.amount)}</td>
                  <td>${formatDate(cash?.entryDate)}</td>
                  <td>${cash?.description || '-'}</td>
                  <td>${match.matchType === 'matched_auto' ? 'Auto' : 'Manual'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
        `}
        
        <div style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd;">
          <p style="font-size: 10px; color: #666;">
            <strong>Report Generated:</strong> ${new Date().toLocaleString()}<br>
            <strong>Application:</strong> ReconPrep v1.0.0<br>
            <strong>Prepared by:</strong> ${preparedByName || '-'}
          </p>
        </div>
      </div>
      
      <div class="footer">
        <span>${company?.companyName} | ${account?.nickname} | ${period?.periodLabel}</span>
        <span>Prepared by: ${preparedByName || '-'} | ReconPrep v1.0.0</span>
      </div>
    </body>
    </html>
  `;
};

// Date formatting
const formatDate = (date) => {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

const formatMoney = (amount, currency = 'MYR') => {
  if (amount === null || amount === undefined) return '-';
  return new Intl.NumberFormat('en-MY', { 
    style: 'currency', 
    currency: currency,
    minimumFractionDigits: 2 
  }).format(amount);
};

// DatePicker Component - defined as a proper React component
function DatePicker({ value, onChange, label, required, style }) {
  const [isOpen, setIsOpen] = React.useState(false);
  const [viewDate, setViewDate] = React.useState(value ? new Date(value) : new Date());
  const containerRef = React.useRef(null);
  
  const selectedDate = value ? new Date(value) : null;
  
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  const days = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
  
  const getDaysInMonth = (year, month) => new Date(year, month + 1, 0).getDate();
  const getFirstDayOfMonth = (year, month) => new Date(year, month, 1).getDay();
  
  const formatDateDisplay = (date) => {
    if (!date) return '';
    const d = new Date(date);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  };
  
  const handleDateClick = (day) => {
    const newDate = new Date(viewDate.getFullYear(), viewDate.getMonth(), day);
    onChange(newDate.toISOString().split('T')[0]);
    setIsOpen(false);
  };
  
  const handlePrevMonth = () => {
    setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() - 1, 1));
  };
  
  const handleNextMonth = () => {
    setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 1));
  };
  
  React.useEffect(() => {
    const handleClickOutside = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);
  
  const daysInMonth = getDaysInMonth(viewDate.getFullYear(), viewDate.getMonth());
  const firstDay = getFirstDayOfMonth(viewDate.getFullYear(), viewDate.getMonth());
  const calendarDays = [];
  
  for (let i = 0; i < firstDay; i++) {
    calendarDays.push(null);
  }
  for (let i = 1; i <= daysInMonth; i++) {
    calendarDays.push(i);
  }
  
  const isSelected = (day) => {
    if (!selectedDate || !day) return false;
    return selectedDate.getFullYear() === viewDate.getFullYear() &&
           selectedDate.getMonth() === viewDate.getMonth() &&
           selectedDate.getDate() === day;
  };
  
  const isToday = (day) => {
    if (!day) return false;
    const today = new Date();
    return today.getFullYear() === viewDate.getFullYear() &&
           today.getMonth() === viewDate.getMonth() &&
           today.getDate() === day;
  };
  
  return (
    <div ref={containerRef} style={{ position: 'relative', ...style }}>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          width: '100%',
          padding: '12px 16px',
          borderRadius: '8px',
          border: '1px solid rgba(148, 163, 184, 0.3)',
          background: 'rgba(15, 23, 42, 0.6)',
          color: value ? '#e2e8f0' : '#64748b',
          fontSize: '14px',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center'
        }}
      >
        <span>{value ? formatDateDisplay(value) : 'Select date...'}</span>
        <span>📅</span>
      </div>
      
      {isOpen && (
        <div style={{
          position: 'absolute',
          top: '100%',
          left: 0,
          marginTop: '4px',
          background: 'rgba(30, 41, 59, 0.98)',
          border: '1px solid rgba(148, 163, 184, 0.3)',
          borderRadius: '12px',
          padding: '16px',
          zIndex: 1000,
          boxShadow: '0 10px 40px rgba(0,0,0,0.5)',
          minWidth: '280px'
        }}>
          {/* Header */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
            <button type="button"
              onClick={handlePrevMonth}
              type="button"
              style={{ background: 'none', border: 'none', color: '#e2e8f0', cursor: 'pointer', fontSize: '18px', padding: '4px 8px' }}
            >
              ‹
            </button>
            <span style={{ fontWeight: '600', color: '#e2e8f0' }}>
              {months[viewDate.getMonth()]} {viewDate.getFullYear()}
            </span>
            <button type="button"
              onClick={handleNextMonth}
              type="button"
              style={{ background: 'none', border: 'none', color: '#e2e8f0', cursor: 'pointer', fontSize: '18px', padding: '4px 8px' }}
            >
              ›
            </button>
          </div>
          
          {/* Day headers */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '2px', marginBottom: '8px' }}>
            {days.map(day => (
              <div key={day} style={{ textAlign: 'center', color: '#64748b', fontSize: '11px', padding: '4px' }}>
                {day}
              </div>
            ))}
          </div>
          
          {/* Calendar grid */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '2px' }}>
            {calendarDays.map((day, idx) => (
              <div
                key={idx}
                onClick={() => day && handleDateClick(day)}
                style={{
                  textAlign: 'center',
                  padding: '8px',
                  borderRadius: '6px',
                  cursor: day ? 'pointer' : 'default',
                  background: isSelected(day) ? 'linear-gradient(135deg, #3b82f6, #8b5cf6)' : 
                              isToday(day) ? 'rgba(59, 130, 246, 0.2)' : 'transparent',
                  color: day ? (isSelected(day) ? 'white' : '#e2e8f0') : 'transparent',
                  fontWeight: isToday(day) || isSelected(day) ? '600' : '400',
                  transition: 'all 0.15s',
                  fontSize: '13px'
                }}
              >
                {day || ''}
              </div>
            ))}
          </div>
          
          {/* Quick actions */}
          <div style={{ display: 'flex', gap: '8px', marginTop: '12px', paddingTop: '12px', borderTop: '1px solid rgba(148, 163, 184, 0.2)' }}>
            <button type="button"
              type="button"
              onClick={() => { onChange(new Date().toISOString().split('T')[0]); setIsOpen(false); }}
              style={{ flex: 1, padding: '6px', background: 'rgba(148, 163, 184, 0.2)', border: 'none', borderRadius: '6px', color: '#e2e8f0', cursor: 'pointer', fontSize: '12px' }}
            >
              Today
            </button>
            <button type="button"
              type="button"
              onClick={() => { onChange(''); setIsOpen(false); }}
              style={{ flex: 1, padding: '6px', background: 'rgba(148, 163, 184, 0.2)', border: 'none', borderRadius: '6px', color: '#94a3b8', cursor: 'pointer', fontSize: '12px' }}
            >
              Clear
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// CSV Parser
const parseCSV = (text) => {
  const lines = text.trim().split('\n');
  if (lines.length < 2) return { headers: [], rows: [] };
  
  const parseRow = (row) => {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < row.length; i++) {
      const char = row[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  };
  
  const headers = parseRow(lines[0]);
  const rows = lines.slice(1).map(line => {
    const values = parseRow(line);
    const obj = {};
    headers.forEach((h, i) => obj[h] = values[i] || '');
    return obj;
  }).filter(row => Object.values(row).some(v => v));
  
  return { headers, rows };
};

// Parse date from various formats
const parseDate = (dateStr) => {
  if (!dateStr) return null;
  const formats = [
    /^(\d{4})-(\d{2})-(\d{2})$/,           // YYYY-MM-DD
    /^(\d{2})\/(\d{2})\/(\d{4})$/,         // DD/MM/YYYY
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,     // D/M/YYYY
  ];
  
  for (const fmt of formats) {
    const match = dateStr.match(fmt);
    if (match) {
      if (fmt === formats[0]) return new Date(match[1], match[2] - 1, match[3]);
      return new Date(match[3], match[2] - 1, match[1]);
    }
  }
  return new Date(dateStr);
};

// ============================================================================
// MAIN APP COMPONENT
// ============================================================================

export default function ReconPrep() {
  // App State
  const [currentScreen, setCurrentScreen] = useState('welcome');
  const [setupStep, setSetupStep] = useState(0);
  const [notification, setNotification] = useState(null);
  const [isInitialized, setIsInitialized] = useState(false);
  
  // Data State - Initialize from localStorage
  // V2: Removed users/preparedByName, added preparedByName
  const [company, setCompany] = useState(null);
  const [preparedByName, setPreparedByName] = useState('');
  const [bankAccounts, setBankAccounts] = useState([]);
  const [selectedAccount, setSelectedAccount] = useState(null);
  const [selectedPeriod, setSelectedPeriod] = useState(null);
  
  // Period Data
  const [periods, setPeriods] = useState({});
  const [bankLines, setBankLines] = useState([]);
  const [cashLines, setCashLines] = useState([]);
  const [matches, setMatches] = useState([]);
  const [notes, setNotes] = useState([]);
  
  // Workspace Tab
  const [workspaceTab, setWorkspaceTab] = useState('imports');
  
  // Matching Config
  const [matchingConfig, setMatchingConfig] = useState({
    dateWindowDays: 3,
    amountTolerance: 0,
    minConfidence: 0.9,
    keywords: ['IBG', 'GIRO', 'TRANSFER', 'DUITNOW', 'CHEQUE', 'CHQ', 'FPX']
  });

  // Load data from localStorage on mount
  // V2: No longer auto-jumps to dashboard - user must explicitly choose
  useEffect(() => {
    // Just mark as initialized, don't auto-load or navigate
    setIsInitialized(true);
  }, []);

  // Save to localStorage when data changes
  // V2: Removed users/currentUser effects, added preparedByName
  useEffect(() => {
    if (!isInitialized) return;
    if (company) saveToStorage(STORAGE_KEYS.COMPANY, company);
  }, [company, isInitialized]);

  useEffect(() => {
    if (!isInitialized) return;
    saveToStorage(STORAGE_KEYS.PREPARED_BY, preparedByName);
  }, [preparedByName, isInitialized]);

  useEffect(() => {
    if (!isInitialized) return;
    saveToStorage(STORAGE_KEYS.BANK_ACCOUNTS, bankAccounts);
  }, [bankAccounts, isInitialized]);

  useEffect(() => {
    if (!isInitialized) return;
    saveToStorage(STORAGE_KEYS.PERIODS, periods);
  }, [periods, isInitialized]);

  useEffect(() => {
    if (!isInitialized) return;
    saveToStorage(STORAGE_KEYS.MATCHING_CONFIG, matchingConfig);
  }, [matchingConfig, isInitialized]);

  // Show notification
  const notify = useCallback((message, type = 'info') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 4000);
  }, []);

  // Generate periods for a bank account
  const generatePeriods = useCallback((account, comp) => {
    const newPeriods = {};
    const startMonth = comp.fyMode === 'fy_start_month' ? comp.fyStartMonth : 
      (new Date(comp.yearEndDate).getMonth() + 2) % 12 || 12;
    const year = new Date().getFullYear();
    
    for (let i = 0; i < 12; i++) {
      const monthIndex = (startMonth - 1 + i) % 12;
      const periodYear = monthIndex < startMonth - 1 ? year + 1 : year;
      const periodId = `${account.id}-${periodYear}-${monthIndex + 1}`;
      const date = new Date(periodYear, monthIndex, 1);
      
      newPeriods[periodId] = {
        id: periodId,
        bankAccountId: account.id,
        monthIndex: i + 1,
        periodLabel: date.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }),
        periodStart: new Date(periodYear, monthIndex, 1).toISOString(),
        periodEnd: new Date(periodYear, monthIndex + 1, 0).toISOString(),
        status: 'not_started',
        bankStatementBalance: null,
        cashBookBalance: null
      };
    }
    return newPeriods;
  }, []);

  // ============================================================================
  // MATCHING ENGINE
  // ============================================================================
  
  const runAutoMatching = useCallback(() => {
    const unmatchedBank = bankLines.filter(b => !matches.find(m => m.bankLineId === b.id));
    const unmatchedCash = cashLines.filter(c => !matches.find(m => m.cashLineId === c.id));
    const newMatches = [];
    const usedBank = new Set();
    const usedCash = new Set();
    
    // Rule 1: Exact amount + same date + reference
    unmatchedBank.forEach(bank => {
      if (usedBank.has(bank.id)) return;
      const candidate = unmatchedCash.find(cash => 
        !usedCash.has(cash.id) &&
        Math.abs(bank.amount - cash.amount) <= matchingConfig.amountTolerance &&
        bank.txnDate === cash.entryDate &&
        bank.reference && cash.reference && 
        bank.reference.toLowerCase() === cash.reference.toLowerCase()
      );
      if (candidate) {
        newMatches.push({
          id: generateId(),
          periodId: selectedPeriod.id,
          bankLineId: bank.id,
          cashLineId: candidate.id,
          matchType: 'matched_auto',
          ruleId: 'R1',
          confidence: 0.98,
          createdAt: new Date().toISOString()
        });
        usedBank.add(bank.id);
        usedCash.add(candidate.id);
      }
    });
    
    // Rule 2: Exact amount + same date
    unmatchedBank.forEach(bank => {
      if (usedBank.has(bank.id)) return;
      const candidate = unmatchedCash.find(cash => 
        !usedCash.has(cash.id) &&
        Math.abs(bank.amount - cash.amount) <= matchingConfig.amountTolerance &&
        bank.txnDate === cash.entryDate
      );
      if (candidate) {
        newMatches.push({
          id: generateId(),
          periodId: selectedPeriod.id,
          bankLineId: bank.id,
          cashLineId: candidate.id,
          matchType: 'matched_auto',
          ruleId: 'R2',
          confidence: 0.95,
          createdAt: new Date().toISOString()
        });
        usedBank.add(bank.id);
        usedCash.add(candidate.id);
      }
    });
    
    // Rule 3: Exact amount + within date window
    unmatchedBank.forEach(bank => {
      if (usedBank.has(bank.id)) return;
      const bankDate = new Date(bank.txnDate);
      const candidate = unmatchedCash.find(cash => {
        if (usedCash.has(cash.id)) return false;
        const cashDate = new Date(cash.entryDate);
        const daysDiff = Math.abs((bankDate - cashDate) / (1000 * 60 * 60 * 24));
        return Math.abs(bank.amount - cash.amount) <= matchingConfig.amountTolerance &&
               daysDiff <= matchingConfig.dateWindowDays;
      });
      if (candidate) {
        newMatches.push({
          id: generateId(),
          periodId: selectedPeriod.id,
          bankLineId: bank.id,
          cashLineId: candidate.id,
          matchType: 'matched_auto',
          ruleId: 'R3',
          confidence: 0.8,
          createdAt: new Date().toISOString()
        });
        usedBank.add(bank.id);
        usedCash.add(candidate.id);
      }
    });
    
    setMatches(prev => [...prev, ...newMatches]);
    notify(`Auto-matching complete: ${newMatches.length} pairs matched`, 'success');
    
    // Update period status
    if (selectedPeriod && selectedPeriod.status === 'not_started') {
      setPeriods(prev => ({
        ...prev,
        [selectedPeriod.id]: { ...prev[selectedPeriod.id], status: 'in_progress' }
      }));
    }
  }, [bankLines, cashLines, matches, matchingConfig, selectedPeriod, notify]);

  // Manual match
  const createManualMatch = useCallback((bankLineId, cashLineId) => {
    const newMatch = {
      id: generateId(),
      periodId: selectedPeriod.id,
      bankLineId,
      cashLineId,
      matchType: 'matched_manual',
      confidence: 1.0,
      createdAt: new Date().toISOString(),
      createdByUserId: null
    };
    setMatches(prev => [...prev, newMatch]);
    notify('Items matched successfully', 'success');
    
    // Check if all items are now matched after this match
    const remainingUnmatchedBank = bankLines.filter(b => 
      b.id !== bankLineId && !matches.find(m => m.bankLineId === b.id)
    );
    const remainingUnmatchedCash = cashLines.filter(c => 
      c.id !== cashLineId && !matches.find(m => m.cashLineId === c.id)
    );
    
    if (remainingUnmatchedBank.length === 0 && remainingUnmatchedCash.length === 0) {
      notify('🎉 All items matched! Navigating to preview...', 'success');
      setTimeout(() => setWorkspaceTab('preview'), 1500);
    }
  }, [selectedPeriod, preparedByName, notify, bankLines, cashLines, matches]);

  // Unmatch
  const removeMatch = useCallback((matchId) => {
    setMatches(prev => prev.filter(m => m.id !== matchId));
    notify('Match removed', 'info');
  }, [notify]);

  // Add or update note
  const addNote = useCallback((lineId, lineType, category, text, existingNoteId = null) => {
    if (existingNoteId) {
      // Update existing note
      setNotes(prev => prev.map(n => 
        n.id === existingNoteId 
          ? { ...n, category, text, updatedAt: new Date().toISOString() }
          : n
      ));
      notify('Note updated', 'success');
    } else {
      // Create new note
      const newNote = {
        id: generateId(),
        periodId: selectedPeriod.id,
        bankLineId: lineType === 'bank' ? lineId : null,
        cashLineId: lineType === 'cash' ? lineId : null,
        category,
        text,
        createdAt: new Date().toISOString(),
        createdByUserId: null
      };
      setNotes(prev => [...prev, newNote]);
      notify('Note added', 'success');
    }
  }, [selectedPeriod, preparedByName, notify]);

  // ============================================================================
  // SAMPLE DATA GENERATORS
  // ============================================================================

  const generateSampleBankStatement = useCallback(() => {
    const sampleData = [
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 1, txnDate: '2024-01-02', description: 'OPENING BALANCE', reference: '', amount: 50000.00, direction: 'inflow', rowHash: 'sample-1', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 2, txnDate: '2024-01-03', description: 'IBG TRANSFER FROM ABC TRADING SDN BHD', reference: 'IBG001234', amount: 15000.00, direction: 'inflow', rowHash: 'sample-2', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 3, txnDate: '2024-01-05', description: 'CHEQUE WITHDRAWAL', reference: 'CHQ-100234', amount: -8500.00, direction: 'outflow', rowHash: 'sample-3', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 4, txnDate: '2024-01-08', description: 'DUITNOW TRANSFER - CUSTOMER PAYMENT', reference: 'DN20240108001', amount: 3200.00, direction: 'inflow', rowHash: 'sample-4', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 5, txnDate: '2024-01-10', description: 'GIRO PAYMENT TO SUPPLIER XYZ', reference: 'GIRO-2024-0055', amount: -12000.00, direction: 'outflow', rowHash: 'sample-5', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 6, txnDate: '2024-01-12', description: 'FPX COLLECTION - ONLINE SALES', reference: 'FPX88776655', amount: 4500.00, direction: 'inflow', rowHash: 'sample-6', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 7, txnDate: '2024-01-15', description: 'BANK SERVICE CHARGE', reference: '', amount: -35.00, direction: 'outflow', rowHash: 'sample-7', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 8, txnDate: '2024-01-15', description: 'INTEREST EARNED', reference: '', amount: 125.50, direction: 'inflow', rowHash: 'sample-8', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 9, txnDate: '2024-01-18', description: 'CHEQUE DEPOSIT - CASH SALES', reference: 'DEP-0012345', amount: 7800.00, direction: 'inflow', rowHash: 'sample-9', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 10, txnDate: '2024-01-20', description: 'SALARY PAYMENT - STAFF', reference: 'SAL-JAN2024', amount: -25000.00, direction: 'outflow', rowHash: 'sample-10', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 11, txnDate: '2024-01-22', description: 'TT RECEIVED FROM OVERSEAS CLIENT', reference: 'TT2024012201', amount: 18500.00, direction: 'inflow', rowHash: 'sample-11', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 12, txnDate: '2024-01-25', description: 'UTILITY BILL PAYMENT - TNB', reference: 'UTIL-TNB-JAN', amount: -1850.00, direction: 'outflow', rowHash: 'sample-12', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 13, txnDate: '2024-01-28', description: 'IBG TRANSFER - RENTAL INCOME', reference: 'IBG009988', amount: 6000.00, direction: 'inflow', rowHash: 'sample-13', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 14, txnDate: '2024-01-30', description: 'CHEQUE WITHDRAWAL - PETTY CASH', reference: 'CHQ-100235', amount: -2000.00, direction: 'outflow', rowHash: 'sample-14', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 15, txnDate: '2024-01-31', description: 'STANDING ORDER - INSURANCE', reference: 'SO-INS-001', amount: -450.00, direction: 'outflow', rowHash: 'sample-15', matchStatus: 'unmatched' },
    ];
    
    setBankLines(sampleData);
    notify(`Loaded 15 sample bank statement lines`, 'success');
    
    if (selectedPeriod?.status === 'not_started') {
      setPeriods(prev => ({
        ...prev,
        [selectedPeriod.id]: { ...prev[selectedPeriod.id], status: 'in_progress' }
      }));
    }
  }, [selectedPeriod, notify]);

  const generateSampleCashBook = useCallback(() => {
    const sampleData = [
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 1, entryDate: '2024-01-02', description: 'Opening Balance B/F', reference: 'OB-JAN', amount: 50000.00, direction: 'inflow', rowHash: 'cb-sample-1', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 2, entryDate: '2024-01-03', description: 'Receipt - ABC Trading Sdn Bhd', reference: 'RCP-001', amount: 15000.00, direction: 'inflow', rowHash: 'cb-sample-2', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 3, entryDate: '2024-01-04', description: 'Payment - Office Supplies', reference: 'CHQ-100234', amount: -8500.00, direction: 'outflow', rowHash: 'cb-sample-3', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 4, entryDate: '2024-01-08', description: 'Receipt - Customer Payment DuitNow', reference: 'RCP-002', amount: 3200.00, direction: 'inflow', rowHash: 'cb-sample-4', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 5, entryDate: '2024-01-10', description: 'Payment - Supplier XYZ', reference: 'PYT-001', amount: -12000.00, direction: 'outflow', rowHash: 'cb-sample-5', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 6, entryDate: '2024-01-12', description: 'Receipt - Online Sales FPX', reference: 'RCP-003', amount: 4500.00, direction: 'inflow', rowHash: 'cb-sample-6', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 7, entryDate: '2024-01-17', description: 'Receipt - Cash Sales Deposit', reference: 'RCP-004', amount: 7800.00, direction: 'inflow', rowHash: 'cb-sample-7', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 8, entryDate: '2024-01-20', description: 'Payment - Staff Salaries January', reference: 'SAL-JAN2024', amount: -25000.00, direction: 'outflow', rowHash: 'cb-sample-8', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 9, entryDate: '2024-01-22', description: 'Receipt - Overseas Client TT', reference: 'RCP-005', amount: 18500.00, direction: 'inflow', rowHash: 'cb-sample-9', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 10, entryDate: '2024-01-25', description: 'Payment - TNB Electricity Bill', reference: 'UTIL-TNB-JAN', amount: -1850.00, direction: 'outflow', rowHash: 'cb-sample-10', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 11, entryDate: '2024-01-28', description: 'Receipt - Rental Income', reference: 'RCP-006', amount: 6000.00, direction: 'inflow', rowHash: 'cb-sample-11', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 12, entryDate: '2024-01-29', description: 'Payment - Petty Cash Replenishment', reference: 'CHQ-100235', amount: -2000.00, direction: 'outflow', rowHash: 'cb-sample-12', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 13, entryDate: '2024-01-31', description: 'Payment - Insurance Premium', reference: 'SO-INS-001', amount: -450.00, direction: 'outflow', rowHash: 'cb-sample-13', matchStatus: 'unmatched' },
      // Items NOT in bank statement (timing differences)
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 14, entryDate: '2024-01-31', description: 'Payment - Supplier DEF (Cheque issued)', reference: 'CHQ-100236', amount: -5500.00, direction: 'outflow', rowHash: 'cb-sample-14', matchStatus: 'unmatched' },
      { id: generateId(), periodId: selectedPeriod?.id, lineNumber: 15, entryDate: '2024-01-31', description: 'Receipt - Customer GHI (Deposit in transit)', reference: 'RCP-007', amount: 9200.00, direction: 'inflow', rowHash: 'cb-sample-15', matchStatus: 'unmatched' },
    ];
    
    setCashLines(sampleData);
    notify(`Loaded 15 sample cash book lines`, 'success');
  }, [selectedPeriod, notify]);

  const downloadSampleCSV = (type) => {
    let csvContent = '';
    let filename = '';
    
    if (type === 'bank') {
      csvContent = `Date,Description,Reference,Amount
2024-01-02,OPENING BALANCE,,50000.00
2024-01-03,IBG TRANSFER FROM ABC TRADING SDN BHD,IBG001234,15000.00
2024-01-05,CHEQUE WITHDRAWAL,CHQ-100234,-8500.00
2024-01-08,DUITNOW TRANSFER - CUSTOMER PAYMENT,DN20240108001,3200.00
2024-01-10,GIRO PAYMENT TO SUPPLIER XYZ,GIRO-2024-0055,-12000.00
2024-01-12,FPX COLLECTION - ONLINE SALES,FPX88776655,4500.00
2024-01-15,BANK SERVICE CHARGE,,-35.00
2024-01-15,INTEREST EARNED,,125.50
2024-01-18,CHEQUE DEPOSIT - CASH SALES,DEP-0012345,7800.00
2024-01-20,SALARY PAYMENT - STAFF,SAL-JAN2024,-25000.00
2024-01-22,TT RECEIVED FROM OVERSEAS CLIENT,TT2024012201,18500.00
2024-01-25,UTILITY BILL PAYMENT - TNB,UTIL-TNB-JAN,-1850.00
2024-01-28,IBG TRANSFER - RENTAL INCOME,IBG009988,6000.00
2024-01-30,CHEQUE WITHDRAWAL - PETTY CASH,CHQ-100235,-2000.00
2024-01-31,STANDING ORDER - INSURANCE,SO-INS-001,-450.00`;
      filename = 'sample_bank_statement.csv';
    } else {
      csvContent = `Date,Description,Reference,Amount
2024-01-02,Opening Balance B/F,OB-JAN,50000.00
2024-01-03,Receipt - ABC Trading Sdn Bhd,RCP-001,15000.00
2024-01-04,Payment - Office Supplies,CHQ-100234,-8500.00
2024-01-08,Receipt - Customer Payment DuitNow,RCP-002,3200.00
2024-01-10,Payment - Supplier XYZ,PYT-001,-12000.00
2024-01-12,Receipt - Online Sales FPX,RCP-003,4500.00
2024-01-17,Receipt - Cash Sales Deposit,RCP-004,7800.00
2024-01-20,Payment - Staff Salaries January,SAL-JAN2024,-25000.00
2024-01-22,Receipt - Overseas Client TT,RCP-005,18500.00
2024-01-25,Payment - TNB Electricity Bill,UTIL-TNB-JAN,-1850.00
2024-01-28,Receipt - Rental Income,RCP-006,6000.00
2024-01-29,Payment - Petty Cash Replenishment,CHQ-100235,-2000.00
2024-01-31,Payment - Insurance Premium,SO-INS-001,-450.00
2024-01-31,Payment - Supplier DEF (Cheque issued),CHQ-100236,-5500.00
2024-01-31,Receipt - Customer GHI (Deposit in transit),RCP-007,9200.00`;
      filename = 'sample_cash_book.csv';
    }
    
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
    notify(`Downloaded ${filename}`, 'success');
  };

  // ============================================================================
  // IMPORT HANDLERS
  // ============================================================================
  
  // FIX C: XLSX SUPPORT - Parse file based on extension
  const parseFileData = useCallback((file) => {
    return new Promise((resolve, reject) => {
      const extension = file.name.split('.').pop().toLowerCase();
      
      if (extension === 'csv') {
        // CSV parsing
        const reader = new FileReader();
        reader.onload = (e) => {
          const { headers, rows } = parseCSV(e.target.result);
          if (!headers.length || !rows.length) {
            reject(new Error('No valid data found in CSV file'));
            return;
          }
          resolve({ headers, rows });
        };
        reader.onerror = () => reject(new Error('Failed to read CSV file'));
        reader.readAsText(file);
      } else if (extension === 'xlsx' || extension === 'xls') {
        // XLSX parsing using SheetJS
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const workbook = XLSX.read(e.target.result, { type: 'array' });
            
            // Use first worksheet only
            const firstSheetName = workbook.SheetNames[0];
            if (!firstSheetName) {
              reject(new Error('No worksheets found in Excel file'));
              return;
            }
            
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
            
            if (!jsonData.length) {
              reject(new Error('No data found in Excel file'));
              return;
            }
            
            // Convert to our format (headers + rows)
            const headers = Object.keys(jsonData[0]);
            const rows = jsonData.map(row => {
              const rowObj = {};
              headers.forEach(h => {
                rowObj[h] = String(row[h] ?? '');
              });
              return rowObj;
            });
            
            resolve({ headers, rows });
          } catch (err) {
            reject(new Error('Failed to parse Excel file: ' + err.message));
          }
        };
        reader.onerror = () => reject(new Error('Failed to read Excel file'));
        reader.readAsArrayBuffer(file);
      } else {
        reject(new Error('Unsupported file format. Please use CSV or XLSX.'));
      }
    });
  }, []);

  const handleBankStatementImport = useCallback(async (file) => {
    try {
      const { headers, rows } = await parseFileData(file);
      
      // Auto-detect columns
      const dateCol = headers.find(h => /date/i.test(h)) || headers[0];
      const descCol = headers.find(h => /desc|narration|particular/i.test(h)) || headers[1];
      const amountCol = headers.find(h => /amount/i.test(h)) || headers[2];
      const refCol = headers.find(h => /ref|cheque|chq/i.test(h));
      
      if (!dateCol || !descCol || !amountCol) {
        notify('Could not detect required columns (Date, Description, Amount)', 'error');
        return;
      }
      
      const newLines = rows.map((row, idx) => {
        const amount = parseFloat(String(row[amountCol])?.replace(/[^0-9.-]/g, '') || 0);
        return {
          id: generateId(),
          periodId: selectedPeriod.id,
          lineNumber: idx + 1,
          txnDate: parseDate(row[dateCol])?.toISOString().split('T')[0] || '',
          description: String(row[descCol] || ''),
          reference: refCol ? String(row[refCol] || '') : '',
          amount: amount,
          direction: amount >= 0 ? 'inflow' : 'outflow',
          rowHash: `${row[dateCol]}-${row[descCol]}-${amount}`,
          matchStatus: 'unmatched'
        };
      }).filter(line => line.txnDate || line.description || line.amount !== 0);
      
      if (newLines.length === 0) {
        notify('No valid data rows found in file', 'error');
        return;
      }
      
      setBankLines(newLines);
      notify(`Imported ${newLines.length} bank statement lines`, 'success');
      
      // Update period status
      if (selectedPeriod.status === 'not_started') {
        setPeriods(prev => ({
          ...prev,
          [selectedPeriod.id]: { ...prev[selectedPeriod.id], status: 'in_progress' }
        }));
      }
    } catch (err) {
      notify(err.message || 'Failed to import file', 'error');
    }
  }, [selectedPeriod, notify, parseFileData]);

  const handleCashBookImport = useCallback(async (file) => {
    try {
      const { headers, rows } = await parseFileData(file);
      
      const dateCol = headers.find(h => /date/i.test(h)) || headers[0];
      const descCol = headers.find(h => /desc|narration|particular/i.test(h)) || headers[1];
      const amountCol = headers.find(h => /amount/i.test(h)) || headers[2];
      const refCol = headers.find(h => /ref|voucher|doc/i.test(h));
      
      if (!dateCol || !descCol || !amountCol) {
        notify('Could not detect required columns (Date, Description, Amount)', 'error');
        return;
      }
      
      const newLines = rows.map((row, idx) => {
        const amount = parseFloat(String(row[amountCol])?.replace(/[^0-9.-]/g, '') || 0);
        return {
          id: generateId(),
          periodId: selectedPeriod.id,
          lineNumber: idx + 1,
          entryDate: parseDate(row[dateCol])?.toISOString().split('T')[0] || '',
          description: String(row[descCol] || ''),
          reference: refCol ? String(row[refCol] || '') : '',
          amount: amount,
          direction: amount >= 0 ? 'inflow' : 'outflow',
          rowHash: `${row[dateCol]}-${row[descCol]}-${amount}`,
          matchStatus: 'unmatched'
        };
      }).filter(line => line.entryDate || line.description || line.amount !== 0);
      
      if (newLines.length === 0) {
        notify('No valid data rows found in file', 'error');
        return;
      }
      
      setCashLines(newLines);
      notify(`Imported ${newLines.length} cash book lines`, 'success');
    } catch (err) {
      notify(err.message || 'Failed to import file', 'error');
    }
  }, [selectedPeriod, notify, parseFileData]);

  // ============================================================================
  // STYLES
  // ============================================================================
  
  const styles = {
    app: {
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%)',
      fontFamily: "'Segoe UI', system-ui, sans-serif",
      color: '#e2e8f0',
      position: 'relative'
    },
    header: {
      background: 'rgba(15, 23, 42, 0.8)',
      backdropFilter: 'blur(12px)',
      borderBottom: '1px solid rgba(148, 163, 184, 0.1)',
      padding: '16px 24px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      position: 'sticky',
      top: 0,
      zIndex: 100
    },
    logo: {
      display: 'flex',
      alignItems: 'center',
      gap: '12px'
    },
    logoIcon: {
      width: '40px',
      height: '40px',
      background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)',
      borderRadius: '10px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '20px',
      fontWeight: 'bold'
    },
    logoText: {
      fontSize: '24px',
      fontWeight: '700',
      background: 'linear-gradient(135deg, #60a5fa, #a78bfa)',
      WebkitBackgroundClip: 'text',
      WebkitTextFillColor: 'transparent'
    },
    container: {
      maxWidth: '1400px',
      margin: '0 auto',
      padding: '24px'
    },
    card: {
      background: 'rgba(30, 41, 59, 0.6)',
      backdropFilter: 'blur(12px)',
      borderRadius: '16px',
      border: '1px solid rgba(148, 163, 184, 0.1)',
      padding: '24px',
      marginBottom: '20px'
    },
    cardTitle: {
      fontSize: '18px',
      fontWeight: '600',
      marginBottom: '16px',
      color: '#f1f5f9'
    },
    btn: {
      padding: '10px 20px',
      borderRadius: '8px',
      border: 'none',
      cursor: 'pointer',
      fontWeight: '500',
      fontSize: '14px',
      transition: 'all 0.2s',
      display: 'inline-flex',
      alignItems: 'center',
      gap: '8px'
    },
    btnPrimary: {
      background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)',
      color: 'white'
    },
    btnSecondary: {
      background: 'rgba(148, 163, 184, 0.2)',
      color: '#e2e8f0',
      border: '1px solid rgba(148, 163, 184, 0.3)'
    },
    btnSuccess: {
      background: 'linear-gradient(135deg, #10b981, #059669)',
      color: 'white'
    },
    btnDanger: {
      background: 'linear-gradient(135deg, #ef4444, #dc2626)',
      color: 'white'
    },
    input: {
      width: '100%',
      padding: '12px 16px',
      borderRadius: '8px',
      border: '1px solid rgba(148, 163, 184, 0.3)',
      background: 'rgba(15, 23, 42, 0.6)',
      color: '#e2e8f0',
      fontSize: '14px',
      outline: 'none',
      transition: 'border-color 0.2s'
    },
    label: {
      display: 'block',
      marginBottom: '6px',
      fontSize: '13px',
      fontWeight: '500',
      color: '#94a3b8'
    },
    formGroup: {
      marginBottom: '16px'
    },
    select: {
      width: '100%',
      padding: '12px 16px',
      borderRadius: '8px',
      border: '1px solid rgba(148, 163, 184, 0.3)',
      background: 'rgba(15, 23, 42, 0.6)',
      color: '#e2e8f0',
      fontSize: '14px',
      outline: 'none'
    },
    table: {
      width: '100%',
      borderCollapse: 'collapse',
      fontSize: '13px'
    },
    th: {
      textAlign: 'left',
      padding: '12px',
      background: 'rgba(15, 23, 42, 0.6)',
      borderBottom: '1px solid rgba(148, 163, 184, 0.2)',
      fontWeight: '600',
      color: '#94a3b8'
    },
    td: {
      padding: '12px',
      borderBottom: '1px solid rgba(148, 163, 184, 0.1)'
    },
    tabs: {
      display: 'flex',
      gap: '4px',
      marginBottom: '20px',
      background: 'rgba(15, 23, 42, 0.4)',
      padding: '4px',
      borderRadius: '12px'
    },
    tab: {
      padding: '10px 20px',
      borderRadius: '8px',
      border: 'none',
      cursor: 'pointer',
      fontWeight: '500',
      fontSize: '14px',
      background: 'transparent',
      color: '#94a3b8',
      transition: 'all 0.2s'
    },
    tabActive: {
      background: 'linear-gradient(135deg, #3b82f6, #8b5cf6)',
      color: 'white'
    },
    grid: {
      display: 'grid',
      gap: '16px'
    },
    notification: {
      position: 'fixed',
      bottom: '24px',
      right: '24px',
      padding: '16px 24px',
      borderRadius: '12px',
      color: 'white',
      fontWeight: '500',
      zIndex: 1000,
      animation: 'slideIn 0.3s ease'
    },
    badge: {
      padding: '4px 10px',
      borderRadius: '12px',
      fontSize: '12px',
      fontWeight: '500'
    },
    statusBadge: {
      not_started: { background: 'rgba(107, 114, 128, 0.3)', color: '#9ca3af' },
      in_progress: { background: 'rgba(245, 158, 11, 0.3)', color: '#fbbf24' },
      pack_generated: { background: 'rgba(16, 185, 129, 0.3)', color: '#34d399' }
    },
    periodGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(4, 1fr)',
      gap: '12px'
    },
    periodCard: {
      padding: '16px',
      borderRadius: '12px',
      border: '1px solid rgba(148, 163, 184, 0.2)',
      cursor: 'pointer',
      transition: 'all 0.2s',
      textAlign: 'center'
    }
  };

  // ============================================================================
  // SCREEN COMPONENTS
  // ============================================================================

  // Welcome Screen
  // V2: Always show both Start New and Load Previous options
  const WelcomeScreen = () => {
    const hasExistingData = loadFromStorage(STORAGE_KEYS.COMPANY) !== null;
    const [showClearConfirm, setShowClearConfirm] = useState(false);
    
    const loadPreviousData = () => {
      const savedCompany = loadFromStorage(STORAGE_KEYS.COMPANY);
      const savedPreparedBy = loadFromStorage(STORAGE_KEYS.PREPARED_BY, '');
      const savedBankAccounts = loadFromStorage(STORAGE_KEYS.BANK_ACCOUNTS, []);
      const savedPeriods = loadFromStorage(STORAGE_KEYS.PERIODS, {});
      const savedMatchingConfig = loadFromStorage(STORAGE_KEYS.MATCHING_CONFIG);
      
      if (!savedCompany || savedBankAccounts.length === 0) {
        notify('No valid saved data found. Please start fresh.', 'error');
        return;
      }
      
      setCompany(savedCompany);
      setPreparedByName(savedPreparedBy);
      setBankAccounts(savedBankAccounts);
      setPeriods(savedPeriods);
      if (savedMatchingConfig) setMatchingConfig(savedMatchingConfig);
      setSelectedAccount(savedBankAccounts[0]);
      setCurrentScreen('dashboard');
      notify('Previous data loaded successfully', 'success');
    };
    
    const clearSavedData = () => {
      Object.values(STORAGE_KEYS).forEach(key => {
        try { localStorage.removeItem(key); } catch (e) {}
      });
      setShowClearConfirm(false);
      notify('All saved data cleared', 'info');
    };
    
    return (
      <div style={{ ...styles.container, maxWidth: '600px', paddingTop: '80px', textAlign: 'center' }}>
        <div style={{ ...styles.logoIcon, width: '80px', height: '80px', fontSize: '36px', margin: '0 auto 24px' }}>
          RP
        </div>
        <h1 style={{ fontSize: '36px', fontWeight: '700', marginBottom: '16px', background: 'linear-gradient(135deg, #60a5fa, #a78bfa)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
          Welcome to ReconPrep
        </h1>
        <p style={{ color: '#94a3b8', fontSize: '16px', marginBottom: '40px', lineHeight: '1.6' }}>
          Professional bank reconciliation preparation tool. Easily import bank statements and cash books, 
          match transactions automatically, and generate professional PDF reconciliation packs.
        </p>
        
        {/* V2: Always show both buttons */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', alignItems: 'center' }}>
          <button type="button" 
            style={{ ...styles.btn, ...styles.btnPrimary, padding: '14px 32px', fontSize: '16px', minWidth: '250px' }}
            onClick={() => {
              // Reset state for fresh start
              setCompany(null);
              setPreparedByName('');
              setBankAccounts([]);
              setPeriods({});
              setSetupStep(0);
              setCurrentScreen('setup');
            }}
          >
            🆕 Start New (Fresh)
          </button>
          
          <button type="button" 
            style={{ 
              ...styles.btn, 
              ...styles.btnSecondary, 
              padding: '14px 32px', 
              fontSize: '16px', 
              minWidth: '250px',
              opacity: hasExistingData ? 1 : 0.5,
              cursor: hasExistingData ? 'pointer' : 'not-allowed'
            }}
            onClick={() => hasExistingData && loadPreviousData()}
            disabled={!hasExistingData}
          >
            📂 Load Previous Data
          </button>
          
          {hasExistingData && (
            <p style={{ color: '#34d399', fontSize: '13px', marginTop: '8px' }}>
              ✓ Saved data available
            </p>
          )}
        </div>
        
        {/* V2: Clear saved data option */}
        {hasExistingData && (
          <div style={{ marginTop: '40px' }}>
            {!showClearConfirm ? (
              <button type="button"
                style={{ background: 'transparent', border: 'none', color: '#64748b', fontSize: '13px', cursor: 'pointer', textDecoration: 'underline' }}
                onClick={() => setShowClearConfirm(true)}
              >
                Clear saved data
              </button>
            ) : (
              <div style={{ 
                background: 'rgba(239, 68, 68, 0.1)', 
                border: '1px solid rgba(239, 68, 68, 0.3)',
                padding: '16px', 
                borderRadius: '12px',
                display: 'inline-block'
              }}>
                <p style={{ color: '#ef4444', marginBottom: '12px', fontSize: '14px' }}>
                  Delete all saved data?
                </p>
                <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
                  <button type="button" 
                    style={{ ...styles.btn, ...styles.btnDanger, padding: '8px 16px', fontSize: '13px' }}
                    onClick={clearSavedData}
                  >
                    Yes, Delete
                  </button>
                  <button type="button" 
                    style={{ ...styles.btn, ...styles.btnSecondary, padding: '8px 16px', fontSize: '13px' }}
                    onClick={() => setShowClearConfirm(false)}
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
        
        <div style={{ marginTop: '60px', color: '#64748b', fontSize: '12px' }}>
          <p>Your data is stored locally in your browser.</p>
          <p style={{ marginTop: '4px' }}>Version 2.0.0</p>
        </div>
      </div>
    );
  };

  // Setup Wizard
  // V2: Removed Users step, added Prepared By field to Company step
  const SetupWizard = () => {
    const [formData, setFormData] = useState({
      companyName: '',
      companyRegNo: '',
      fyMode: 'year_end_date',
      yearEndDate: '',
      fyStartMonth: 1,
      currency: 'MYR',
      preparedBy: ''
    });
    const [newAccount, setNewAccount] = useState({ bankName: '', nickname: '', accountRef: '', openingBalance: '' });

    // V2: Only 2 steps now
    const steps = ['Company Profile', 'Bank Accounts'];

    const handleCompanySave = () => {
      if (!formData.companyName) {
        notify('Please enter company name', 'error');
        return;
      }
      setCompany({ 
        companyName: formData.companyName,
        companyRegNo: formData.companyRegNo,
        fyMode: formData.fyMode,
        yearEndDate: formData.yearEndDate,
        fyStartMonth: formData.fyStartMonth,
        currency: formData.currency,
        id: generateId(), 
        createdAt: new Date().toISOString() 
      });
      setPreparedByName(formData.preparedBy);
      setSetupStep(1);
      notify('Company profile saved', 'success');
    };

    const handleAddAccount = () => {
      if (!newAccount.bankName || !newAccount.nickname) {
        notify('Please fill in bank name and nickname', 'error');
        return;
      }
      const account = { 
        ...newAccount, 
        id: generateId(), 
        companyId: company.id,
        openingBalance: newAccount.openingBalance ? parseFloat(newAccount.openingBalance) : null,
        isActive: true,
        createdAt: new Date().toISOString() 
      };
      setBankAccounts(prev => [...prev, account]);
      
      // Generate periods for this account
      const newPeriods = generatePeriods(account, company);
      setPeriods(prev => ({ ...prev, ...newPeriods }));
      
      setNewAccount({ bankName: '', nickname: '', accountRef: '', openingBalance: '' });
      notify('Bank account added', 'success');
    };

    const finishSetup = () => {
      if (bankAccounts.length === 0) {
        notify('Please add at least one bank account', 'error');
        return;
      }
      setSelectedAccount(bankAccounts[0]);
      setCurrentScreen('dashboard');
      notify('Setup complete!', 'success');
    };

    return (
      <div style={{ ...styles.container, maxWidth: '700px', paddingTop: '40px' }}>
        {/* Progress - V2: 2 steps only */}
        <div style={{ display: 'flex', justifyContent: 'center', gap: '8px', marginBottom: '40px' }}>
          {steps.map((step, i) => (
            <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div style={{
                width: '32px', height: '32px', borderRadius: '50%',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '14px', fontWeight: '600',
                background: i <= setupStep ? 'linear-gradient(135deg, #3b82f6, #8b5cf6)' : 'rgba(148, 163, 184, 0.2)',
                color: i <= setupStep ? 'white' : '#64748b'
              }}>
                {i + 1}
              </div>
              <span style={{ color: i <= setupStep ? '#e2e8f0' : '#64748b', fontSize: '14px' }}>{step}</span>
              {i < steps.length - 1 && <div style={{ width: '40px', height: '2px', background: i < setupStep ? '#3b82f6' : 'rgba(148, 163, 184, 0.2)' }} />}
            </div>
          ))}
        </div>

        {/* Step 0: Company Profile + Prepared By */}
        {setupStep === 0 && (
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>Company Profile</h2>
            <div style={styles.formGroup}>
              <label style={styles.label}>Company Name *</label>
              <input 
                style={styles.input}
                placeholder="Enter company name"
                value={formData.companyName}
                onChange={e => setFormData({ ...formData, companyName: e.target.value })}
              />
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Registration No.</label>
              <input 
                style={styles.input}
                placeholder="e.g., 201901012345"
                value={formData.companyRegNo}
                onChange={e => setFormData({ ...formData, companyRegNo: e.target.value })}
              />
            </div>
            <div style={styles.formGroup}>
              <label style={styles.label}>Financial Year Definition *</label>
              <select 
                style={styles.select}
                value={formData.fyMode}
                onChange={e => setFormData({ ...formData, fyMode: e.target.value })}
              >
                <option value="year_end_date">Year End Date</option>
                <option value="fy_start_month">FY Start Month</option>
              </select>
            </div>
            {formData.fyMode === 'year_end_date' ? (
              <div style={styles.formGroup}>
                <label style={styles.label}>Year End Date *</label>
                <DatePicker
                  value={formData.yearEndDate}
                  onChange={(date) => setFormData({ ...formData, yearEndDate: date })}
                />
              </div>
            ) : (
              <div style={styles.formGroup}>
                <label style={styles.label}>FY Start Month *</label>
                <select 
                  style={styles.select}
                  value={formData.fyStartMonth}
                  onChange={e => setFormData({ ...formData, fyStartMonth: parseInt(e.target.value) })}
                >
                  {['January','February','March','April','May','June','July','August','September','October','November','December'].map((m, i) => (
                    <option key={i} value={i + 1}>{m}</option>
                  ))}
                </select>
              </div>
            )}
            <div style={styles.formGroup}>
              <label style={styles.label}>Currency *</label>
              <select 
                style={styles.select}
                value={formData.currency}
                onChange={e => setFormData({ ...formData, currency: e.target.value })}
              >
                <option value="MYR">MYR - Malaysian Ringgit</option>
                <option value="USD">USD - US Dollar</option>
                <option value="SGD">SGD - Singapore Dollar</option>
                <option value="GBP">GBP - British Pound</option>
                <option value="EUR">EUR - Euro</option>
              </select>
            </div>
            
            {/* V2: Prepared By field */}
            <div style={styles.formGroup}>
              <label style={styles.label}>Prepared By (Name)</label>
              <input 
                style={styles.input}
                placeholder="Your name (optional, for PDF header)"
                value={formData.preparedBy}
                onChange={e => setFormData({ ...formData, preparedBy: e.target.value })}
              />
              <p style={{ color: '#64748b', fontSize: '12px', marginTop: '4px' }}>
                This name will appear in generated PDF reports.
              </p>
            </div>
            
            <button type="button" style={{ ...styles.btn, ...styles.btnPrimary }} onClick={handleCompanySave}>
              Continue →
            </button>
          </div>
        )}

        {/* Step 1: Bank Accounts (was Step 2) */}
        {setupStep === 1 && (
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>Bank Accounts</h2>
            
            {bankAccounts.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={styles.th}>Bank</th>
                      <th style={styles.th}>Nickname</th>
                      <th style={styles.th}>Opening Balance</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bankAccounts.map(a => (
                      <tr key={a.id}>
                        <td style={styles.td}>{a.bankName}</td>
                        <td style={styles.td}>{a.nickname}</td>
                        <td style={styles.td}>{a.openingBalance ? formatMoney(a.openingBalance, company?.currency) : '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '16px' }}>
              <div>
                <label style={styles.label}>Bank Name *</label>
                <input 
                  style={styles.input}
                  placeholder="e.g., Maybank"
                  value={newAccount.bankName}
                  onChange={e => setNewAccount({ ...newAccount, bankName: e.target.value })}
                />
              </div>
              <div>
                <label style={styles.label}>Account Nickname *</label>
                <input 
                  style={styles.input}
                  placeholder="e.g., Main Operating"
                  value={newAccount.nickname}
                  onChange={e => setNewAccount({ ...newAccount, nickname: e.target.value })}
                />
              </div>
              <div>
                <label style={styles.label}>Account Ref</label>
                <input 
                  style={styles.input}
                  placeholder="Optional reference"
                  value={newAccount.accountRef}
                  onChange={e => setNewAccount({ ...newAccount, accountRef: e.target.value })}
                />
              </div>
              <div>
                <label style={styles.label}>Opening Balance</label>
                <input 
                  type="number"
                  step="0.01"
                  style={styles.input}
                  placeholder="0.00"
                  value={newAccount.openingBalance}
                  onChange={e => setNewAccount({ ...newAccount, openingBalance: e.target.value })}
                />
              </div>
            </div>
            <div style={{ display: 'flex', gap: '12px' }}>
              <button type="button" style={{ ...styles.btn, ...styles.btnSecondary }} onClick={handleAddAccount}>
                + Add Account
              </button>
              <button type="button" style={{ ...styles.btn, ...styles.btnPrimary }} onClick={finishSetup}>
                Finish Setup ✓
              </button>
            </div>
          </div>
        )}
      </div>
    );
  };

  // Dashboard
  const Dashboard = () => {
    const accountPeriods = selectedAccount ? 
      Object.values(periods).filter(p => p.bankAccountId === selectedAccount.id) : [];

    return (
      <div style={styles.container}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
          <div>
            <h1 style={{ fontSize: '28px', fontWeight: '700', marginBottom: '4px' }}>{company?.companyName}</h1>
            <p style={{ color: '#94a3b8' }}>Financial Year {company?.fyMode === 'year_end_date' ? `ending ${formatDate(company?.yearEndDate)}` : `starting ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][company?.fyStartMonth - 1]}`}</p>
          </div>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <select 
              style={{ ...styles.select, width: 'auto' }}
              value={selectedAccount?.id || ''}
              onChange={e => setSelectedAccount(bankAccounts.find(a => a.id === e.target.value))}
            >
              {bankAccounts.map(a => (
                <option key={a.id} value={a.id}>{a.bankName} - {a.nickname}</option>
              ))}
            </select>
            <button type="button" style={{ ...styles.btn, ...styles.btnSecondary }} onClick={() => setCurrentScreen('settings')}>
              ⚙ Settings
            </button>
          </div>
        </div>

        {/* Period Grid */}
        <div style={styles.card}>
          <h3 style={styles.cardTitle}>Select Period</h3>
          <div style={styles.periodGrid}>
            {accountPeriods.sort((a, b) => a.monthIndex - b.monthIndex).map(period => (
              <div 
                key={period.id}
                style={{
                  ...styles.periodCard,
                  background: selectedPeriod?.id === period.id ? 'rgba(59, 130, 246, 0.2)' : 'rgba(15, 23, 42, 0.4)',
                  borderColor: selectedPeriod?.id === period.id ? '#3b82f6' : 'rgba(148, 163, 184, 0.2)'
                }}
                onClick={() => setSelectedPeriod(period)}
              >
                <div style={{ fontSize: '14px', fontWeight: '600', marginBottom: '8px' }}>{period.periodLabel}</div>
                <span style={{ ...styles.badge, ...styles.statusBadge[period.status] }}>
                  {period.status.replace(/_/g, ' ')}
                </span>
              </div>
            ))}
          </div>
          {selectedPeriod && (
            <div style={{ marginTop: '20px', textAlign: 'center' }}>
              <button type="button" 
                style={{ ...styles.btn, ...styles.btnPrimary, padding: '14px 32px' }}
                onClick={() => setCurrentScreen('workspace')}
              >
                Open {selectedPeriod.periodLabel} →
              </button>
            </div>
          )}
        </div>

        {/* Quick Stats */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
          <div style={styles.card}>
            <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '4px' }}>Bank Accounts</div>
            <div style={{ fontSize: '28px', fontWeight: '700' }}>{bankAccounts.length}</div>
          </div>
          <div style={styles.card}>
            <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '4px' }}>In Progress</div>
            <div style={{ fontSize: '28px', fontWeight: '700', color: '#fbbf24' }}>
              {Object.values(periods).filter(p => p.status === 'in_progress').length}
            </div>
          </div>
          <div style={styles.card}>
            <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '4px' }}>Packs Generated</div>
            <div style={{ fontSize: '28px', fontWeight: '700', color: '#34d399' }}>
              {Object.values(periods).filter(p => p.status === 'pack_generated').length}
            </div>
          </div>
          <div style={styles.card}>
            <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '4px' }}>Prepared By</div>
            <div style={{ fontSize: '16px', fontWeight: '600' }}>{preparedByName || '-'}</div>
          </div>
        </div>
      </div>
    );
  };

  // Workspace
  const Workspace = () => {
    const [selectedBankLines, setSelectedBankLines] = useState([]);
    const [selectedCashLines, setSelectedCashLines] = useState([]);
    const [noteModal, setNoteModal] = useState(null);
    const [newNote, setNewNote] = useState({ category: 'timing_difference', text: '' });
    const [confirmedItems, setConfirmedItems] = useState({ bank: [], cash: [] });

    // Items that are unmatched AND not confirmed as reconciling items
    const unmatchedBank = bankLines.filter(b => 
      !matches.find(m => m.bankLineId === b.id) && 
      !confirmedItems.bank.includes(b.id)
    );
    const unmatchedCash = cashLines.filter(c => 
      !matches.find(m => m.cashLineId === c.id) && 
      !confirmedItems.cash.includes(c.id)
    );
    
    // Confirmed reconciling items (have notes and confirmed by user)
    const confirmedBankItems = bankLines.filter(b => confirmedItems.bank.includes(b.id));
    const confirmedCashItems = cashLines.filter(c => confirmedItems.cash.includes(c.id));
    
    const matchedPairs = matches.map(m => ({
      match: m,
      bank: bankLines.find(b => b.id === m.bankLineId),
      cash: cashLines.find(c => c.id === m.cashLineId)
    }));

    const bankTotal = bankLines.reduce((sum, l) => sum + l.amount, 0);
    const cashTotal = cashLines.reduce((sum, l) => sum + l.amount, 0);
    const unmatchedBankTotal = unmatchedBank.reduce((sum, l) => sum + l.amount, 0);
    const unmatchedCashTotal = unmatchedCash.reduce((sum, l) => sum + l.amount, 0);
    const confirmedBankTotal = confirmedBankItems.reduce((sum, l) => sum + l.amount, 0);
    const confirmedCashTotal = confirmedCashItems.reduce((sum, l) => sum + l.amount, 0);

    // Confirm item as reconciling item (must have a note)
    const confirmAsReconcilingItem = (lineId, type) => {
      const hasNote = type === 'bank' 
        ? notes.find(n => n.bankLineId === lineId)
        : notes.find(n => n.cashLineId === lineId);
      
      if (!hasNote) {
        notify('Please add a note before confirming this item', 'error');
        return;
      }
      
      if (type === 'bank') {
        setConfirmedItems(prev => ({ ...prev, bank: [...prev.bank, lineId] }));
      } else {
        setConfirmedItems(prev => ({ ...prev, cash: [...prev.cash, lineId] }));
      }
      notify('Item confirmed as reconciling item', 'success');
      
      // Clear selection
      setSelectedBankLines([]);
      setSelectedCashLines([]);
    };
    
    // Unconfirm item (move back to unmatched)
    const unconfirmItem = (lineId, type) => {
      if (type === 'bank') {
        setConfirmedItems(prev => ({ ...prev, bank: prev.bank.filter(id => id !== lineId) }));
      } else {
        setConfirmedItems(prev => ({ ...prev, cash: prev.cash.filter(id => id !== lineId) }));
      }
      notify('Item moved back to unmatched', 'info');
    };

    // Confirm all items that have notes
    const confirmAllWithNotes = () => {
      const bankWithNotes = unmatchedBank.filter(b => notes.find(n => n.bankLineId === b.id)).map(b => b.id);
      const cashWithNotes = unmatchedCash.filter(c => notes.find(n => n.cashLineId === c.id)).map(c => c.id);
      
      if (bankWithNotes.length === 0 && cashWithNotes.length === 0) {
        notify('No items with notes to confirm', 'error');
        return;
      }
      
      setConfirmedItems(prev => ({
        bank: [...prev.bank, ...bankWithNotes],
        cash: [...prev.cash, ...cashWithNotes]
      }));
      notify(`Confirmed ${bankWithNotes.length + cashWithNotes.length} items as reconciling items`, 'success');
    };

    const handleManualMatch = () => {
      if (selectedBankLines.length !== 1 || selectedCashLines.length !== 1) {
        notify('Select exactly one bank line and one cash book line', 'error');
        return;
      }
      createManualMatch(selectedBankLines[0], selectedCashLines[0]);
      setSelectedBankLines([]);
      setSelectedCashLines([]);
    };

    const handleSaveNote = () => {
      if (!noteModal) return;
      addNote(
        noteModal.lineId, 
        noteModal.type, 
        newNote.category, 
        newNote.text,
        noteModal.existingNote?.id || null
      );
      setNoteModal(null);
      setNewNote({ category: 'timing_difference', text: '' });
    };

    // Initialize note form when opening modal for editing
    React.useEffect(() => {
      if (noteModal?.existingNote) {
        setNewNote({
          category: noteModal.existingNote.category,
          text: noteModal.existingNote.text || ''
        });
      } else if (noteModal) {
        setNewNote({ category: 'timing_difference', text: '' });
      }
    }, [noteModal]);

    const generateReport = async () => {
      // FIX A: PDF PRINT VIEW - Open in new window with auto-print
      const pdfContent = generatePDFContent({
        company,
        account: selectedAccount,
        period: selectedPeriod,
        bankLines,
        cashLines,
        matches,
        notes,
        preparedByName,
        confirmedItems,
        stats: {
          bankTotal,
          cashTotal,
          unmatchedBankTotal,
          unmatchedCashTotal,
          confirmedBankTotal,
          confirmedCashTotal,
          matchRate: bankLines.length > 0 ? Math.round((matches.length / bankLines.length) * 100) : 0
        }
      });
      
      // Open new window for print view
      const printWindow = window.open('', '_blank');
      
      if (printWindow) {
        // Write content to new window
        printWindow.document.write(pdfContent);
        printWindow.document.close();
        
        // Auto-trigger print dialog after content loads
        printWindow.onload = function() {
          setTimeout(function() {
            printWindow.print();
          }, 500);
        };
        
        // Fallback if onload doesn't fire (some browsers)
        setTimeout(function() {
          try {
            printWindow.print();
          } catch (e) {
            // Print already triggered or window closed
          }
        }, 1000);
        
        notify('Report opened in new tab. Use Print → Save as PDF to save.', 'success');
      } else {
        // Popup blocked
        notify('Popup blocked! Please allow popups for this site to generate the report.', 'error');
        return; // Don't update status if failed
      }
      
      // Update period status
      setPeriods(prev => ({
        ...prev,
        [selectedPeriod.id]: { ...prev[selectedPeriod.id], status: 'pack_generated' }
      }));
    };
    
    // FIX A: Removed pdfPreviewContent state and openPreviewInNewTab as no longer needed

    return (
      <div style={styles.container}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            <button type="button" 
              style={{ ...styles.btn, ...styles.btnSecondary, padding: '8px 12px' }}
              onClick={() => setCurrentScreen('dashboard')}
            >
              ← Back
            </button>
            <div>
              <h1 style={{ fontSize: '24px', fontWeight: '700', marginBottom: '4px' }}>
                {selectedAccount?.bankName} - {selectedAccount?.nickname}
              </h1>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <span style={{ color: '#94a3b8' }}>{selectedPeriod?.periodLabel}</span>
                <span style={{ ...styles.badge, ...styles.statusBadge[selectedPeriod?.status] }}>
                  {selectedPeriod?.status?.replace(/_/g, ' ')}
                </span>
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '12px' }}>
            <button type="button" style={{ ...styles.btn, ...styles.btnSecondary }} onClick={(e) => { e.preventDefault(); runAutoMatching(); }}>
              ⚡ Run Auto-Match
            </button>
            <button type="button" style={{ ...styles.btn, ...styles.btnSuccess }} onClick={(e) => { e.preventDefault(); generateReport(); }}>
              📄 Generate PDF Pack
            </button>
          </div>
        </div>

        {/* Tabs */}
        <div style={styles.tabs}>
          {['imports', 'matching', 'unmatched', 'preview'].map(tab => (
            <button type="button"
              key={tab}
              style={{ ...styles.tab, ...(workspaceTab === tab ? styles.tabActive : {}) }}
              onClick={() => setWorkspaceTab(tab)}
            >
              {tab === 'imports' && '📥 '}
              {tab === 'matching' && '🔗 '}
              {tab === 'unmatched' && '📋 '}
              {tab === 'preview' && '👁 '}
              {tab.charAt(0).toUpperCase() + tab.slice(1)}
            </button>
          ))}
        </div>

        {/* Imports Tab */}
        {workspaceTab === 'imports' && (
          <div>
            {/* Help Banner */}
            <div style={{
              background: 'rgba(59, 130, 246, 0.1)',
              border: '1px solid rgba(59, 130, 246, 0.3)',
              borderRadius: '12px',
              padding: '16px',
              marginBottom: '20px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start'
            }}>
              <span style={{ fontSize: '20px' }}>💡</span>
              <div>
                <strong style={{ color: '#60a5fa' }}>Step 1: Import Your Data</strong>
                <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px', lineHeight: '1.5' }}>
                  Upload your <strong>Bank Statement</strong> and <strong>Cash Book</strong> files (CSV format). 
                  The system will automatically detect columns like Date, Description, Amount, and Reference. 
                  You can also use sample data to test the workflow.
                </p>
              </div>
            </div>

            {/* Quick Load Sample Data Banner */}
            <div style={{ 
              ...styles.card, 
              background: 'linear-gradient(135deg, rgba(59, 130, 246, 0.2), rgba(139, 92, 246, 0.2))',
              border: '1px solid rgba(99, 102, 241, 0.3)',
              marginBottom: '20px'
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: '16px', marginBottom: '4px' }}>🚀 Quick Start with Sample Data</h3>
                  <p style={{ margin: 0, color: '#94a3b8', fontSize: '13px' }}>
                    Load pre-built sample data to test the reconciliation workflow instantly
                  </p>
                </div>
                <div style={{ display: 'flex', gap: '12px' }}>
                  <button type="button" 
                    style={{ ...styles.btn, ...styles.btnPrimary }}
                    onClick={() => { generateSampleBankStatement(); generateSampleCashBook(); }}
                  >
                    ⚡ Load All Sample Data
                  </button>
                </div>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              <div style={styles.card}>
                <h3 style={styles.cardTitle}>📄 Bank Statement</h3>
                <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '16px' }}>
                  Import CSV file with Date, Description, Amount columns
                </p>
                
                {/* File Upload */}
                <div style={{ marginBottom: '16px' }}>
                  <input 
                    type="file"
                    accept=".csv"
                    onChange={e => e.target.files[0] && handleBankStatementImport(e.target.files[0])}
                    style={{ marginBottom: '12px' }}
                  />
                </div>

                {/* Sample Data Options */}
                <div style={{ 
                  background: 'rgba(15, 23, 42, 0.4)', 
                  padding: '12px', 
                  borderRadius: '8px',
                  marginBottom: '16px'
                }}>
                  <div style={{ fontSize: '12px', color: '#94a3b8', marginBottom: '8px' }}>Or use sample data:</div>
                  <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <button type="button" 
                      style={{ ...styles.btn, ...styles.btnSecondary, padding: '6px 12px', fontSize: '12px' }}
                      onClick={generateSampleBankStatement}
                    >
                      📥 Load Sample
                    </button>
                    <button type="button" 
                      style={{ ...styles.btn, padding: '6px 12px', fontSize: '12px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                      onClick={() => downloadSampleCSV('bank')}
                    >
                      ⬇ Download CSV
                    </button>
                  </div>
                </div>

                {/* Status */}
                {bankLines.length > 0 ? (
                  <div style={{ background: 'rgba(16, 185, 129, 0.1)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.3)' }}>
                    <strong style={{ color: '#34d399' }}>✓ {bankLines.length} lines loaded</strong>
                    <div style={{ color: '#94a3b8', fontSize: '13px', marginTop: '4px' }}>
                      Total: {formatMoney(bankTotal, company?.currency)}
                    </div>
                    <button type="button" 
                      style={{ ...styles.btn, padding: '4px 8px', fontSize: '11px', marginTop: '8px', background: 'transparent', border: '1px solid rgba(239, 68, 68, 0.5)', color: '#ef4444' }}
                      onClick={() => setBankLines([])}
                    >
                      Clear Data
                    </button>
                  </div>
                ) : (
                  <div style={{ background: 'rgba(107, 114, 128, 0.1)', padding: '12px', borderRadius: '8px', textAlign: 'center', color: '#94a3b8' }}>
                    No data loaded yet
                  </div>
                )}
              </div>

              <div style={styles.card}>
                <h3 style={styles.cardTitle}>📒 Cash Book</h3>
                <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '16px' }}>
                  Import CSV or XLSX file with Date, Description, Amount columns
                </p>
                
                {/* File Upload */}
                <div style={{ marginBottom: '16px' }}>
                  <input 
                    type="file"
                    accept=".csv,.xlsx"
                    onChange={e => e.target.files[0] && handleCashBookImport(e.target.files[0])}
                    style={{ marginBottom: '12px' }}
                  />
                </div>

                {/* Sample Data Options */}
                <div style={{ 
                  background: 'rgba(15, 23, 42, 0.4)', 
                  padding: '12px', 
                  borderRadius: '8px',
                  marginBottom: '16px'
                }}>
                  <div style={{ fontSize: '12px', color: '#94a3b8', marginBottom: '8px' }}>Or use sample data:</div>
                  <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    <button type="button" 
                      style={{ ...styles.btn, ...styles.btnSecondary, padding: '6px 12px', fontSize: '12px' }}
                      onClick={generateSampleCashBook}
                    >
                      📥 Load Sample
                    </button>
                    <button type="button" 
                      style={{ ...styles.btn, padding: '6px 12px', fontSize: '12px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                      onClick={() => downloadSampleCSV('cash')}
                    >
                      ⬇ Download CSV
                    </button>
                  </div>
                </div>

                {/* Status */}
                {cashLines.length > 0 ? (
                  <div style={{ background: 'rgba(16, 185, 129, 0.1)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.3)' }}>
                    <strong style={{ color: '#34d399' }}>✓ {cashLines.length} lines loaded</strong>
                    <div style={{ color: '#94a3b8', fontSize: '13px', marginTop: '4px' }}>
                      Total: {formatMoney(cashTotal, company?.currency)}
                    </div>
                    <button type="button" 
                      style={{ ...styles.btn, padding: '4px 8px', fontSize: '11px', marginTop: '8px', background: 'transparent', border: '1px solid rgba(239, 68, 68, 0.5)', color: '#ef4444' }}
                      onClick={() => setCashLines([])}
                    >
                      Clear Data
                    </button>
                  </div>
                ) : (
                  <div style={{ background: 'rgba(107, 114, 128, 0.1)', padding: '12px', borderRadius: '8px', textAlign: 'center', color: '#94a3b8' }}>
                    No data loaded yet
                  </div>
                )}
              </div>
            </div>

            {/* Data Preview Tables */}
            {(bankLines.length > 0 || cashLines.length > 0) && (
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginTop: '20px' }}>
                {bankLines.length > 0 && (
                  <div style={styles.card}>
                    <h3 style={styles.cardTitle}>Bank Statement Preview (First 10 rows)</h3>
                    <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
                      <table style={styles.table}>
                        <thead>
                          <tr>
                            <th style={styles.th}>#</th>
                            <th style={styles.th}>Date</th>
                            <th style={styles.th}>Description</th>
                            <th style={styles.th}>Amount</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bankLines.slice(0, 10).map((line, idx) => (
                            <tr key={line.id}>
                              <td style={styles.td}>{idx + 1}</td>
                              <td style={styles.td}>{formatDate(line.txnDate)}</td>
                              <td style={{ ...styles.td, maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {line.description}
                              </td>
                              <td style={{ 
                                ...styles.td, 
                                textAlign: 'right', 
                                fontFamily: 'monospace',
                                color: line.amount >= 0 ? '#34d399' : '#ef4444'
                              }}>
                                {formatMoney(line.amount, company?.currency)}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {bankLines.length > 10 && (
                        <div style={{ textAlign: 'center', padding: '8px', color: '#94a3b8', fontSize: '12px' }}>
                          ... and {bankLines.length - 10} more rows
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {cashLines.length > 0 && (
                  <div style={styles.card}>
                    <h3 style={styles.cardTitle}>Cash Book Preview (First 10 rows)</h3>
                    <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
                      <table style={styles.table}>
                        <thead>
                          <tr>
                            <th style={styles.th}>#</th>
                            <th style={styles.th}>Date</th>
                            <th style={styles.th}>Description</th>
                            <th style={styles.th}>Amount</th>
                          </tr>
                        </thead>
                        <tbody>
                          {cashLines.slice(0, 10).map((line, idx) => (
                            <tr key={line.id}>
                              <td style={styles.td}>{idx + 1}</td>
                              <td style={styles.td}>{formatDate(line.entryDate)}</td>
                              <td style={{ ...styles.td, maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {line.description}
                              </td>
                              <td style={{ 
                                ...styles.td, 
                                textAlign: 'right', 
                                fontFamily: 'monospace',
                                color: line.amount >= 0 ? '#34d399' : '#ef4444'
                              }}>
                                {formatMoney(line.amount, company?.currency)}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {cashLines.length > 10 && (
                        <div style={{ textAlign: 'center', padding: '8px', color: '#94a3b8', fontSize: '12px' }}>
                          ... and {cashLines.length - 10} more rows
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Next Step Prompt */}
            {bankLines.length > 0 && cashLines.length > 0 && (
              <div style={{ 
                ...styles.card, 
                background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.1), rgba(5, 150, 105, 0.1))',
                border: '1px solid rgba(16, 185, 129, 0.3)',
                marginTop: '20px',
                textAlign: 'center'
              }}>
                <h3 style={{ margin: 0, marginBottom: '8px', color: '#34d399' }}>✓ Data Ready!</h3>
                <p style={{ margin: 0, color: '#94a3b8', marginBottom: '16px' }}>
                  Both bank statement and cash book loaded. Ready to run matching.
                </p>
                <button type="button" 
                  style={{ ...styles.btn, ...styles.btnSuccess, padding: '12px 24px' }}
                  onClick={() => { runAutoMatching(); setWorkspaceTab('matching'); }}
                >
                  ⚡ Run Auto-Match & View Results
                </button>
              </div>
            )}
          </div>
        )}

        {/* Matching Tab */}
        {workspaceTab === 'matching' && (
          <div>
            {/* Help Banner */}
            <div style={{
              background: 'rgba(59, 130, 246, 0.1)',
              border: '1px solid rgba(59, 130, 246, 0.3)',
              borderRadius: '12px',
              padding: '16px',
              marginBottom: '20px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start'
            }}>
              <span style={{ fontSize: '20px' }}>💡</span>
              <div>
                <strong style={{ color: '#60a5fa' }}>Step 2: Review Matched Items</strong>
                <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px', lineHeight: '1.5' }}>
                  Click <strong>"Run Auto-Match"</strong> to automatically pair bank statement entries with cash book entries 
                  based on amount, date, and reference. Review the matched pairs below. If any match is incorrect, 
                  click <strong>"Unmatch"</strong> to separate them. Then go to the <strong>Unmatched & Notes</strong> tab 
                  to handle remaining items.
                </p>
              </div>
            </div>

            {/* Summary */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '16px', marginBottom: '20px' }}>
              <div style={{ ...styles.card, padding: '16px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Bank Lines</div>
                <div style={{ fontSize: '24px', fontWeight: '700' }}>{bankLines.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '16px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Cash Lines</div>
                <div style={{ fontSize: '24px', fontWeight: '700' }}>{cashLines.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '16px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Matched</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#34d399' }}>{matches.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '16px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Unmatched Bank</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#ef4444' }}>{unmatchedBank.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '16px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Unmatched Cash</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#ef4444' }}>{unmatchedCash.length}</div>
              </div>
            </div>

            {/* Matched Pairs Table */}
            <div style={styles.card}>
              <h3 style={styles.cardTitle}>Matched Pairs ({matches.length})</h3>
              {matchedPairs.length === 0 ? (
                <p style={{ color: '#94a3b8', textAlign: 'center', padding: '40px' }}>
                  No matches yet. Import data and run auto-match.
                </p>
              ) : (
                <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={styles.th}>Bank Date</th>
                        <th style={styles.th}>Bank Description</th>
                        <th style={styles.th}>Amount</th>
                        <th style={styles.th}>Cash Date</th>
                        <th style={styles.th}>Cash Description</th>
                        <th style={styles.th}>Type</th>
                        <th style={styles.th}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {matchedPairs.map(({ match, bank, cash }) => (
                        <tr key={match.id}>
                          <td style={styles.td}>{formatDate(bank?.txnDate)}</td>
                          <td style={{ ...styles.td, maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {bank?.description}
                          </td>
                          <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace' }}>
                            {formatMoney(bank?.amount, company?.currency)}
                          </td>
                          <td style={styles.td}>{formatDate(cash?.entryDate)}</td>
                          <td style={{ ...styles.td, maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {cash?.description}
                          </td>
                          <td style={styles.td}>
                            <span style={{ ...styles.badge, background: match.matchType === 'matched_auto' ? 'rgba(16, 185, 129, 0.2)' : 'rgba(59, 130, 246, 0.2)', color: match.matchType === 'matched_auto' ? '#34d399' : '#60a5fa' }}>
                              {match.matchType === 'matched_auto' ? 'Auto' : 'Manual'}
                            </span>
                          </td>
                          <td style={styles.td}>
                            <button type="button" 
                              style={{ ...styles.btn, ...styles.btnDanger, padding: '4px 8px', fontSize: '12px' }}
                              onClick={() => removeMatch(match.id)}
                            >
                              Unmatch
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Unmatched Tab */}
        {workspaceTab === 'unmatched' && (
          <div>
            {/* Help Banner */}
            <div style={{
              background: 'rgba(59, 130, 246, 0.1)',
              border: '1px solid rgba(59, 130, 246, 0.3)',
              borderRadius: '12px',
              padding: '16px',
              marginBottom: '20px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start'
            }}>
              <span style={{ fontSize: '20px' }}>💡</span>
              <div>
                <strong style={{ color: '#60a5fa' }}>Step 3: Handle Unmatched Items</strong>
                <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px', lineHeight: '1.5' }}>
                  <strong>Option A - Manual Match:</strong> If two items should be paired, select ONE from each side and click "Match Selected".<br/>
                  <strong>Option B - Reconciling Items:</strong> For items that can't be matched (bank charges, timing differences), 
                  add a <strong>Note</strong> to categorize them, then click <strong>"✓ Confirm"</strong> to mark them as reconciling items.
                </p>
              </div>
            </div>

            {/* Summary Stats */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '12px', marginBottom: '20px' }}>
              <div style={{ ...styles.card, padding: '12px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '11px' }}>Pending Bank</div>
                <div style={{ fontSize: '20px', fontWeight: '700', color: '#ef4444' }}>{unmatchedBank.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '12px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '11px' }}>Pending Cash</div>
                <div style={{ fontSize: '20px', fontWeight: '700', color: '#ef4444' }}>{unmatchedCash.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '12px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '11px' }}>Confirmed Bank</div>
                <div style={{ fontSize: '20px', fontWeight: '700', color: '#8b5cf6' }}>{confirmedBankItems.length}</div>
              </div>
              <div style={{ ...styles.card, padding: '12px', textAlign: 'center' }}>
                <div style={{ color: '#94a3b8', fontSize: '11px' }}>Confirmed Cash</div>
                <div style={{ fontSize: '20px', fontWeight: '700', color: '#8b5cf6' }}>{confirmedCashItems.length}</div>
              </div>
            </div>

            {/* Info about why items might not match */}
            {unmatchedBank.length > 0 && unmatchedCash.length > 0 && (
              <div style={{
                background: 'rgba(245, 158, 11, 0.1)',
                border: '1px solid rgba(245, 158, 11, 0.3)',
                borderRadius: '12px',
                padding: '16px',
                marginBottom: '20px',
                display: 'flex',
                gap: '12px',
                alignItems: 'flex-start'
              }}>
                <span style={{ fontSize: '20px' }}>⚠️</span>
                <div>
                  <strong style={{ color: '#fbbf24' }}>Can't find a match?</strong>
                  <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px', lineHeight: '1.5' }}>
                    Not all items can be matched! Common reasons:<br/>
                    • <strong>Bank charges/interest</strong> - appear only in bank statement, not yet recorded in cash book<br/>
                    • <strong>Outstanding cheques</strong> - recorded in cash book but not yet cleared by bank<br/>
                    • <strong>Deposits in transit</strong> - recorded in cash book but not yet credited by bank<br/>
                    <br/>
                    <em>Add notes to categorize these items instead of matching them.</em>
                  </p>
                </div>
              </div>
            )}

            {/* Manual Match Controls */}
            <div style={{ ...styles.card, marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                <div>
                  <strong>Actions</strong>
                  <div style={{ color: '#94a3b8', fontSize: '13px', marginTop: '4px' }}>
                    {selectedBankLines.length === 0 && selectedCashLines.length === 0 ? (
                      <span>Select items to match, or add notes and click ✓ to confirm individually</span>
                    ) : selectedBankLines.length === 1 && selectedCashLines.length === 0 ? (
                      <span>✓ Bank item selected. Select a cash item to match them together</span>
                    ) : selectedBankLines.length === 0 && selectedCashLines.length === 1 ? (
                      <span>✓ Cash item selected. Select a bank item to match them together</span>
                    ) : selectedBankLines.length === 1 && selectedCashLines.length === 1 ? (
                      <span style={{ color: '#34d399' }}>✓ Ready! Click "Match Selected" to pair these items</span>
                    ) : (
                      <span style={{ color: '#fbbf24' }}>⚠ Select only ONE item from each side to match</span>
                    )}
                  </div>
                </div>
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                  <button type="button" 
                    style={{ ...styles.btn, ...styles.btnSecondary, padding: '8px 12px', fontSize: '12px' }}
                    onClick={() => { 
                      setSelectedBankLines([]); 
                      setSelectedCashLines([]); 
                    }}
                  >
                    Clear
                  </button>
                  <button type="button" 
                    style={{ 
                      ...styles.btn, 
                      ...styles.btnPrimary,
                      padding: '8px 12px', 
                      fontSize: '12px',
                      opacity: (selectedBankLines.length !== 1 || selectedCashLines.length !== 1) ? 0.5 : 1,
                      cursor: (selectedBankLines.length !== 1 || selectedCashLines.length !== 1) ? 'not-allowed' : 'pointer'
                    }}
                    onClick={() => {
                      if (selectedBankLines.length === 1 && selectedCashLines.length === 1) {
                        handleManualMatch();
                      }
                    }}
                    title="Select one bank item AND one cash item to match them as the same transaction"
                  >
                    🔗 Match Selected
                  </button>
                  <button type="button" 
                    style={{ 
                      ...styles.btn, 
                      background: 'rgba(139, 92, 246, 0.2)',
                      border: '1px solid rgba(139, 92, 246, 0.5)',
                      color: '#a78bfa',
                      padding: '8px 12px', 
                      fontSize: '12px'
                    }}
                    onClick={confirmAllWithNotes}
                    title="Confirm all items that have notes as reconciling items"
                  >
                    ✓✓ Confirm All with Notes
                  </button>
                </div>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              {/* Unmatched Bank */}
              <div style={styles.card}>
                <h3 style={styles.cardTitle}>Pending Bank Items ({unmatchedBank.length})</h3>
                <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '12px' }}>
                  Total: {formatMoney(unmatchedBankTotal, company?.currency)}
                </div>
                {unmatchedBank.length === 0 ? (
                  <div style={{ background: 'rgba(16, 185, 129, 0.1)', padding: '20px', borderRadius: '8px', textAlign: 'center' }}>
                    <span style={{ color: '#34d399' }}>✓ All bank items processed!</span>
                  </div>
                ) : (
                  <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                    <table style={styles.table}>
                      <thead>
                        <tr>
                          <th style={{ ...styles.th, width: '30px' }}></th>
                          <th style={styles.th}>Date</th>
                          <th style={styles.th}>Description</th>
                          <th style={styles.th}>Amount</th>
                          <th style={styles.th}>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {unmatchedBank.map(line => {
                          const lineNote = notes.find(n => n.bankLineId === line.id);
                          return (
                            <tr key={line.id} style={{ background: selectedBankLines.includes(line.id) ? 'rgba(59, 130, 246, 0.2)' : 'transparent' }}>
                              <td style={styles.td}>
                                <input 
                                  type="checkbox"
                                  checked={selectedBankLines.includes(line.id)}
                                  onChange={e => {
                                    if (e.target.checked) {
                                      setSelectedBankLines(prev => [...prev, line.id]);
                                    } else {
                                      setSelectedBankLines(prev => prev.filter(id => id !== line.id));
                                    }
                                  }}
                                />
                              </td>
                              <td style={styles.td}>{formatDate(line.txnDate)}</td>
                              <td style={{ ...styles.td, maxWidth: '100px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={line.description}>
                                {line.description}
                              </td>
                              <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', color: line.amount >= 0 ? '#34d399' : '#ef4444', fontSize: '12px' }}>
                                {formatMoney(line.amount, company?.currency)}
                              </td>
                              <td style={{ ...styles.td, whiteSpace: 'nowrap' }}>
                                <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                                  {lineNote ? (
                                    <button type="button"
                                      style={{ 
                                        ...styles.badge, 
                                        background: 'rgba(139, 92, 246, 0.2)', 
                                        color: '#a78bfa',
                                        border: 'none',
                                        cursor: 'pointer',
                                        fontSize: '10px',
                                        padding: '2px 6px'
                                      }}
                                      onClick={() => setNoteModal({ lineId: line.id, type: 'bank', existingNote: lineNote })}
                                      title="Click to edit note"
                                    >
                                      ✎ {lineNote.category.replace(/_/g, ' ').substring(0, 10)}
                                    </button>
                                  ) : (
                                    <button type="button" 
                                      style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                                      onClick={() => setNoteModal({ lineId: line.id, type: 'bank', existingNote: null })}
                                    >
                                      + Note
                                    </button>
                                  )}
                                  {lineNote && (
                                    <button type="button" 
                                      style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'rgba(16, 185, 129, 0.2)', border: '1px solid rgba(16, 185, 129, 0.5)', color: '#34d399' }}
                                      onClick={() => confirmAsReconcilingItem(line.id, 'bank')}
                                      title="Confirm as reconciling item"
                                    >
                                      ✓
                                    </button>
                                  )}
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {/* Unmatched Cash */}
              <div style={styles.card}>
                <h3 style={styles.cardTitle}>Unmatched Cash Items ({unmatchedCash.length})</h3>
                <div style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '12px' }}>
                  Total: {formatMoney(unmatchedCashTotal, company?.currency)}
                </div>
                {unmatchedCash.length === 0 ? (
                  <div style={{ background: 'rgba(16, 185, 129, 0.1)', padding: '20px', borderRadius: '8px', textAlign: 'center' }}>
                    <span style={{ color: '#34d399' }}>✓ All cash book items processed!</span>
                  </div>
                ) : (
                  <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                    <table style={styles.table}>
                      <thead>
                        <tr>
                          <th style={{ ...styles.th, width: '30px' }}></th>
                          <th style={styles.th}>Date</th>
                          <th style={styles.th}>Description</th>
                          <th style={styles.th}>Amount</th>
                          <th style={styles.th}>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {unmatchedCash.map(line => {
                          const lineNote = notes.find(n => n.cashLineId === line.id);
                          return (
                            <tr key={line.id} style={{ background: selectedCashLines.includes(line.id) ? 'rgba(59, 130, 246, 0.2)' : 'transparent' }}>
                              <td style={styles.td}>
                                <input 
                                  type="checkbox"
                                  checked={selectedCashLines.includes(line.id)}
                                  onChange={e => {
                                    if (e.target.checked) {
                                      setSelectedCashLines(prev => [...prev, line.id]);
                                    } else {
                                      setSelectedCashLines(prev => prev.filter(id => id !== line.id));
                                    }
                                  }}
                                />
                              </td>
                              <td style={styles.td}>{formatDate(line.entryDate)}</td>
                              <td style={{ ...styles.td, maxWidth: '100px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={line.description}>
                                {line.description}
                              </td>
                              <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', color: line.amount >= 0 ? '#34d399' : '#ef4444', fontSize: '12px' }}>
                                {formatMoney(line.amount, company?.currency)}
                              </td>
                              <td style={{ ...styles.td, whiteSpace: 'nowrap' }}>
                                <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                                  {lineNote ? (
                                    <button type="button"
                                      style={{ 
                                        ...styles.badge, 
                                        background: 'rgba(139, 92, 246, 0.2)', 
                                        color: '#a78bfa',
                                        border: 'none',
                                        cursor: 'pointer',
                                        fontSize: '10px',
                                        padding: '2px 6px'
                                      }}
                                      onClick={() => setNoteModal({ lineId: line.id, type: 'cash', existingNote: lineNote })}
                                      title="Click to edit note"
                                    >
                                      ✎ {lineNote.category.replace(/_/g, ' ').substring(0, 10)}
                                    </button>
                                  ) : (
                                    <button type="button" 
                                      style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                                      onClick={() => setNoteModal({ lineId: line.id, type: 'cash', existingNote: null })}
                                    >
                                      + Note
                                    </button>
                                  )}
                                  {lineNote && (
                                    <button type="button" 
                                      style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'rgba(16, 185, 129, 0.2)', border: '1px solid rgba(16, 185, 129, 0.5)', color: '#34d399' }}
                                      onClick={() => confirmAsReconcilingItem(line.id, 'cash')}
                                      title="Confirm as reconciling item"
                                    >
                                      ✓
                                    </button>
                                  )}
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>

            {/* Confirmed Reconciling Items */}
            {(confirmedBankItems.length > 0 || confirmedCashItems.length > 0) && (
              <div style={{ ...styles.card, marginTop: '20px', background: 'rgba(139, 92, 246, 0.05)', border: '1px solid rgba(139, 92, 246, 0.2)' }}>
                <h3 style={{ ...styles.cardTitle, color: '#a78bfa' }}>✓ Confirmed Reconciling Items</h3>
                <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '16px' }}>
                  These items have been reviewed and confirmed as reconciling differences.
                </p>
                
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                  {/* Confirmed Bank Items */}
                  {confirmedBankItems.length > 0 && (
                    <div>
                      <h4 style={{ fontSize: '13px', color: '#94a3b8', marginBottom: '8px' }}>
                        Bank Items ({confirmedBankItems.length}) - Total: {formatMoney(confirmedBankTotal, company?.currency)}
                      </h4>
                      <div style={{ maxHeight: '200px', overflowY: 'auto' }}>
                        {confirmedBankItems.map(line => {
                          const lineNote = notes.find(n => n.bankLineId === line.id);
                          return (
                            <div key={line.id} style={{ 
                              display: 'flex', 
                              justifyContent: 'space-between', 
                              alignItems: 'center',
                              padding: '8px',
                              background: 'rgba(15, 23, 42, 0.4)',
                              borderRadius: '6px',
                              marginBottom: '4px',
                              fontSize: '12px'
                            }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ color: '#e2e8f0' }}>{line.description.substring(0, 30)}...</div>
                                <div style={{ color: '#8b5cf6', fontSize: '10px' }}>{lineNote?.category.replace(/_/g, ' ')}</div>
                              </div>
                              <div style={{ textAlign: 'right', marginRight: '8px' }}>
                                <span style={{ fontFamily: 'monospace', color: line.amount >= 0 ? '#34d399' : '#ef4444' }}>
                                  {formatMoney(line.amount, company?.currency)}
                                </span>
                              </div>
                              <button type="button"
                                style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                                onClick={() => unconfirmItem(line.id, 'bank')}
                                title="Move back to pending"
                              >
                                ↩
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  
                  {/* Confirmed Cash Items */}
                  {confirmedCashItems.length > 0 && (
                    <div>
                      <h4 style={{ fontSize: '13px', color: '#94a3b8', marginBottom: '8px' }}>
                        Cash Book Items ({confirmedCashItems.length}) - Total: {formatMoney(confirmedCashTotal, company?.currency)}
                      </h4>
                      <div style={{ maxHeight: '200px', overflowY: 'auto' }}>
                        {confirmedCashItems.map(line => {
                          const lineNote = notes.find(n => n.cashLineId === line.id);
                          return (
                            <div key={line.id} style={{ 
                              display: 'flex', 
                              justifyContent: 'space-between', 
                              alignItems: 'center',
                              padding: '8px',
                              background: 'rgba(15, 23, 42, 0.4)',
                              borderRadius: '6px',
                              marginBottom: '4px',
                              fontSize: '12px'
                            }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ color: '#e2e8f0' }}>{line.description.substring(0, 30)}...</div>
                                <div style={{ color: '#8b5cf6', fontSize: '10px' }}>{lineNote?.category.replace(/_/g, ' ')}</div>
                              </div>
                              <div style={{ textAlign: 'right', marginRight: '8px' }}>
                                <span style={{ fontFamily: 'monospace', color: line.amount >= 0 ? '#34d399' : '#ef4444' }}>
                                  {formatMoney(line.amount, company?.currency)}
                                </span>
                              </div>
                              <button type="button"
                                style={{ ...styles.btn, padding: '2px 6px', fontSize: '10px', background: 'transparent', border: '1px solid rgba(148, 163, 184, 0.3)', color: '#94a3b8' }}
                                onClick={() => unconfirmItem(line.id, 'cash')}
                                title="Move back to pending"
                              >
                                ↩
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Completion Check */}
            {unmatchedBank.length === 0 && unmatchedCash.length === 0 && (
              <div style={{ 
                ...styles.card, 
                background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.1), rgba(5, 150, 105, 0.1))',
                border: '1px solid rgba(16, 185, 129, 0.3)',
                marginTop: '20px',
                textAlign: 'center'
              }}>
                <h3 style={{ margin: 0, marginBottom: '8px', color: '#34d399' }}>🎉 All Items Processed!</h3>
                <p style={{ margin: 0, color: '#94a3b8', marginBottom: '16px' }}>
                  All bank statement and cash book items have been matched or confirmed as reconciling items.
                </p>
                <button type="button" 
                  style={{ ...styles.btn, ...styles.btnSuccess, padding: '12px 24px' }}
                  onClick={() => setWorkspaceTab('preview')}
                >
                  👁 View Report Preview
                </button>
              </div>
            )}

            {/* Note Modal */}
            {noteModal && (
              <div style={{
                position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
                background: 'rgba(0,0,0,0.7)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
              }}>
                <div style={{ ...styles.card, width: '450px', margin: 0 }}>
                  <h3 style={styles.cardTitle}>
                    {noteModal.existingNote ? '✎ Edit Note' : '+ Add Note'}
                  </h3>
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Category</label>
                    <select 
                      style={styles.select}
                      value={newNote.category}
                      onChange={e => setNewNote({ ...newNote, category: e.target.value })}
                    >
                      <option value="timing_difference">Timing Difference</option>
                      <option value="missing_cashbook_entry">Missing Cash Book Entry</option>
                      <option value="bank_charge_interest">Bank Charge / Interest</option>
                      <option value="cheque_not_presented">Cheque Not Presented</option>
                      <option value="deposit_in_transit">Deposit in Transit</option>
                      <option value="error_correction">Error / Correction Required</option>
                      <option value="unknown">Unknown / To Investigate</option>
                    </select>
                  </div>
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Note (optional)</label>
                    <textarea 
                      style={{ ...styles.input, minHeight: '80px', resize: 'vertical' }}
                      placeholder="Add details..."
                      value={newNote.text}
                      onChange={e => setNewNote({ ...newNote, text: e.target.value })}
                    />
                  </div>
                  <div style={{ display: 'flex', gap: '12px', justifyContent: 'space-between' }}>
                    <div>
                      {noteModal.existingNote && (
                        <button type="button" 
                          style={{ ...styles.btn, ...styles.btnDanger }}
                          onClick={() => {
                            setNotes(prev => prev.filter(n => n.id !== noteModal.existingNote.id));
                            setNoteModal(null);
                            setNewNote({ category: 'timing_difference', text: '' });
                            notify('Note deleted', 'info');
                          }}
                        >
                          🗑 Delete Note
                        </button>
                      )}
                    </div>
                    <div style={{ display: 'flex', gap: '12px' }}>
                      <button type="button" 
                        style={{ ...styles.btn, ...styles.btnSecondary }} 
                        onClick={() => {
                          setNoteModal(null);
                          setNewNote({ category: 'timing_difference', text: '' });
                        }}
                      >
                        Cancel
                      </button>
                      <button type="button" style={{ ...styles.btn, ...styles.btnPrimary }} onClick={handleSaveNote}>
                        {noteModal.existingNote ? 'Update Note' : 'Save Note'}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Preview Tab */}
        {workspaceTab === 'preview' && (
          <div>
            {/* Help Banner */}
            <div style={{
              background: 'rgba(59, 130, 246, 0.1)',
              border: '1px solid rgba(59, 130, 246, 0.3)',
              borderRadius: '12px',
              padding: '16px',
              marginBottom: '20px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start'
            }}>
              <span style={{ fontSize: '20px' }}>💡</span>
              <div>
                <strong style={{ color: '#60a5fa' }}>Step 4: Review & Generate Report</strong>
                <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px', lineHeight: '1.5' }}>
                  Review the reconciliation summary below. When ready, click <strong>"Generate PDF Pack"</strong> to create 
                  a professional reconciliation report. A new window will open with the report - use your browser's 
                  <strong> Print → Save as PDF</strong> option to save it.
                </p>
              </div>
            </div>

            <div style={styles.card}>
            <h3 style={styles.cardTitle}>Reconciliation Summary Preview</h3>
            
            <div style={{ background: 'rgba(15, 23, 42, 0.6)', padding: '24px', borderRadius: '12px', marginBottom: '20px' }}>
              <div style={{ textAlign: 'center', marginBottom: '24px' }}>
                <h2 style={{ margin: 0, fontSize: '20px' }}>{company?.companyName}</h2>
                <p style={{ color: '#94a3b8', margin: '4px 0' }}>Bank Reconciliation Statement</p>
                <p style={{ color: '#94a3b8', margin: 0 }}>
                  {selectedAccount?.bankName} - {selectedAccount?.nickname} | {selectedPeriod?.periodLabel}
                </p>
              </div>

              <table style={{ ...styles.table, maxWidth: '500px', margin: '0 auto' }}>
                <tbody>
                  <tr>
                    <td style={{ ...styles.td, fontWeight: '600' }}>Balance per Bank Statement</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace' }}>
                      {formatMoney(bankTotal, company?.currency)}
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...styles.td, paddingLeft: '24px', color: '#94a3b8' }}>Less: Outstanding Cheques</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', color: '#ef4444' }}>
                      {formatMoney(unmatchedCash.filter(l => l.amount < 0).reduce((s, l) => s + l.amount, 0), company?.currency)}
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...styles.td, paddingLeft: '24px', color: '#94a3b8' }}>Add: Deposits in Transit</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', color: '#34d399' }}>
                      {formatMoney(unmatchedCash.filter(l => l.amount > 0).reduce((s, l) => s + l.amount, 0), company?.currency)}
                    </td>
                  </tr>
                  <tr style={{ borderTop: '2px solid rgba(148, 163, 184, 0.3)' }}>
                    <td style={{ ...styles.td, fontWeight: '700' }}>Adjusted Bank Balance</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', fontWeight: '700' }}>
                      {formatMoney(bankTotal + unmatchedCashTotal, company?.currency)}
                    </td>
                  </tr>
                  <tr><td colSpan={2} style={{ height: '16px' }}></td></tr>
                  <tr>
                    <td style={{ ...styles.td, fontWeight: '600' }}>Balance per Cash Book</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace' }}>
                      {formatMoney(cashTotal, company?.currency)}
                    </td>
                  </tr>
                  <tr>
                    <td style={{ ...styles.td, paddingLeft: '24px', color: '#94a3b8' }}>Add/Less: Unrecorded Items</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace' }}>
                      {formatMoney(unmatchedBankTotal, company?.currency)}
                    </td>
                  </tr>
                  <tr style={{ borderTop: '2px solid rgba(148, 163, 184, 0.3)' }}>
                    <td style={{ ...styles.td, fontWeight: '700' }}>Adjusted Cash Book Balance</td>
                    <td style={{ ...styles.td, textAlign: 'right', fontFamily: 'monospace', fontWeight: '700' }}>
                      {formatMoney(cashTotal + unmatchedBankTotal, company?.currency)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Statistics */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
              <div style={{ background: 'rgba(15, 23, 42, 0.4)', padding: '16px', borderRadius: '8px' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Bank Lines</div>
                <div style={{ fontSize: '24px', fontWeight: '700' }}>{bankLines.length}</div>
              </div>
              <div style={{ background: 'rgba(15, 23, 42, 0.4)', padding: '16px', borderRadius: '8px' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Cash Lines</div>
                <div style={{ fontSize: '24px', fontWeight: '700' }}>{cashLines.length}</div>
              </div>
              <div style={{ background: 'rgba(15, 23, 42, 0.4)', padding: '16px', borderRadius: '8px' }}>
                <div style={{ color: '#94a3b8', fontSize: '12px' }}>Matched Pairs</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#34d399' }}>{matches.length}</div>
              </div>
              <div style={{ background: 'rgba(139, 92, 246, 0.1)', padding: '16px', borderRadius: '8px', border: '1px solid rgba(139, 92, 246, 0.3)' }}>
                <div style={{ color: '#a78bfa', fontSize: '12px' }}>Confirmed Reconciling</div>
                <div style={{ fontSize: '24px', fontWeight: '700', color: '#a78bfa' }}>{confirmedBankItems.length + confirmedCashItems.length}</div>
              </div>
            </div>
            
            {/* Pending Warning */}
            {(unmatchedBank.length > 0 || unmatchedCash.length > 0) && (
              <div style={{ 
                background: 'rgba(245, 158, 11, 0.1)', 
                border: '1px solid rgba(245, 158, 11, 0.3)',
                borderRadius: '8px',
                padding: '12px',
                marginTop: '16px',
                display: 'flex',
                alignItems: 'center',
                gap: '12px'
              }}>
                <span style={{ fontSize: '20px' }}>⚠️</span>
                <div>
                  <strong style={{ color: '#fbbf24' }}>Pending Items</strong>
                  <p style={{ margin: '4px 0 0 0', color: '#94a3b8', fontSize: '13px' }}>
                    You have {unmatchedBank.length} bank and {unmatchedCash.length} cash book items still pending. 
                    Go to "Unmatched & Notes" tab to add notes and confirm them.
                  </p>
                </div>
              </div>
            )}

            <div style={{ marginTop: '24px', textAlign: 'center' }}>
              <button type="button" 
                type="button"
                style={{ ...styles.btn, ...styles.btnSuccess, padding: '14px 32px' }} 
                onClick={(e) => { e.preventDefault(); generateReport(); }}
              >
                📄 Generate PDF Pack
              </button>
              <p style={{ color: '#94a3b8', fontSize: '13px', marginTop: '12px' }}>
                A new tab will open with the report. Use Print → Save as PDF to save.
              </p>
              <p style={{ color: '#64748b', fontSize: '12px', marginTop: '8px' }}>
                Prepared by: {preparedByName} | {new Date().toLocaleString()}
              </p>
            </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  // Settings Screen
  const Settings = () => {
    const [showClearConfirm, setShowClearConfirm] = useState(false);
    
    const clearAllData = () => {
      Object.values(STORAGE_KEYS).forEach(key => {
        try { localStorage.removeItem(key); } catch (e) {}
      });
      setCompany(null);
      setPreparedByName('');
      setBankAccounts([]);
      setPeriods({});
      setCurrentScreen('welcome');
      notify('All data cleared. Starting fresh.', 'info');
    };
    
    return (
    <div style={styles.container}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '24px' }}>
        <button type="button" 
          style={{ ...styles.btn, ...styles.btnSecondary, padding: '8px 12px' }}
          onClick={() => setCurrentScreen('dashboard')}
        >
          ← Back
        </button>
        <h1 style={{ fontSize: '24px', fontWeight: '700', margin: 0 }}>Settings</h1>
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>Company Profile</h3>
        <p><strong>Name:</strong> {company?.companyName}</p>
        <p><strong>Registration:</strong> {company?.companyRegNo || '-'}</p>
        <p><strong>Currency:</strong> {company?.currency}</p>
      </div>

      {/* V2: Prepared By field instead of Users table */}
      <div style={styles.card}>
        <h3 style={styles.cardTitle}>Prepared By</h3>
        <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '12px' }}>
          This name appears in generated PDF reports.
        </p>
        <input 
          style={{ ...styles.input, maxWidth: '300px' }}
          placeholder="Enter preparer name"
          value={preparedByName}
          onChange={e => setPreparedByName(e.target.value)}
        />
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>Bank Accounts ({bankAccounts.length})</h3>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>Bank</th>
              <th style={styles.th}>Nickname</th>
              <th style={styles.th}>Opening Balance</th>
            </tr>
          </thead>
          <tbody>
            {bankAccounts.map(a => (
              <tr key={a.id}>
                <td style={styles.td}>{a.bankName}</td>
                <td style={styles.td}>{a.nickname}</td>
                <td style={styles.td}>{formatMoney(a.openingBalance, company?.currency)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>Matching Configuration</h3>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '16px' }}>
          <div>
            <label style={styles.label}>Date Window (Days)</label>
            <input 
              type="number"
              style={styles.input}
              value={matchingConfig.dateWindowDays}
              onChange={e => setMatchingConfig({ ...matchingConfig, dateWindowDays: parseInt(e.target.value) })}
            />
          </div>
          <div>
            <label style={styles.label}>Amount Tolerance</label>
            <input 
              type="number"
              step="0.01"
              style={styles.input}
              value={matchingConfig.amountTolerance}
              onChange={e => setMatchingConfig({ ...matchingConfig, amountTolerance: parseFloat(e.target.value) })}
            />
          </div>
          <div>
            <label style={styles.label}>Min Confidence</label>
            <input 
              type="number"
              step="0.1"
              min="0"
              max="1"
              style={styles.input}
              value={matchingConfig.minConfidence}
              onChange={e => setMatchingConfig({ ...matchingConfig, minConfidence: parseFloat(e.target.value) })}
            />
          </div>
        </div>
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>Data Management</h3>
        <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '16px' }}>
          Your data is stored locally in your browser. Clearing data will remove all company, user, and bank account information.
        </p>
        {!showClearConfirm ? (
          <button type="button" 
            style={{ ...styles.btn, ...styles.btnDanger }}
            onClick={() => setShowClearConfirm(true)}
          >
            🗑 Clear All Data
          </button>
        ) : (
          <div style={{ background: 'rgba(239, 68, 68, 0.1)', padding: '16px', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.3)' }}>
            <p style={{ color: '#ef4444', marginBottom: '12px', fontWeight: '500' }}>
              ⚠️ Are you sure? This will permanently delete all your data.
            </p>
            <div style={{ display: 'flex', gap: '12px' }}>
              <button type="button" 
                style={{ ...styles.btn, ...styles.btnDanger }}
                onClick={clearAllData}
              >
                Yes, Clear Everything
              </button>
              <button type="button" 
                style={{ ...styles.btn, ...styles.btnSecondary }}
                onClick={() => setShowClearConfirm(false)}
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>About</h3>
        <p><strong>ReconPrep</strong> v2.0.0</p>
        <p style={{ color: '#94a3b8' }}>Bank Reconciliation Preparation Tool</p>
        <p style={{ color: '#64748b', fontSize: '12px', marginTop: '8px' }}>
          Data is stored locally in your browser's localStorage.
        </p>
      </div>
    </div>
  );
  };

  // ============================================================================
  // RENDER
  // ============================================================================

  return (
    <div style={styles.app}>
      {/* Header */}
      {currentScreen !== 'welcome' && (
        <header style={styles.header}>
          <div style={styles.logo}>
            <div style={styles.logoIcon}>RP</div>
            <span style={styles.logoText}>ReconPrep</span>
          </div>
          {/* V2: Show preparedByName instead of currentUser */}
          {preparedByName && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
              <span style={{ color: '#94a3b8' }}>👤 {preparedByName}</span>
            </div>
          )}
        </header>
      )}

      {/* Screens */}
      {currentScreen === 'welcome' && <WelcomeScreen />}
      {currentScreen === 'setup' && <SetupWizard />}
      {currentScreen === 'dashboard' && <Dashboard />}
      {currentScreen === 'workspace' && <Workspace />}
      {currentScreen === 'settings' && <Settings />}

      {/* Notification */}
      {notification && (
        <div style={{
          ...styles.notification,
          background: notification.type === 'success' ? 'linear-gradient(135deg, #10b981, #059669)' :
                      notification.type === 'error' ? 'linear-gradient(135deg, #ef4444, #dc2626)' :
                      'linear-gradient(135deg, #3b82f6, #8b5cf6)'
        }}>
          {notification.message}
        </div>
      )}

      <style>{`
        @keyframes slideIn {
          from { transform: translateX(100%); opacity: 0; }
          to { transform: translateX(0); opacity: 1; }
        }
        input[type="file"] {
          color: #94a3b8;
        }
        input[type="file"]::file-selector-button {
          background: linear-gradient(135deg, #3b82f6, #8b5cf6);
          color: white;
          border: none;
          padding: 8px 16px;
          border-radius: 6px;
          cursor: pointer;
          margin-right: 12px;
        }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: rgba(15, 23, 42, 0.4); border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: rgba(148, 163, 184, 0.3); border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: rgba(148, 163, 184, 0.5); }
      `}</style>
    </div>
  );
}