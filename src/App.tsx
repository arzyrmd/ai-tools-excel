import React, { useState, useEffect, useRef } from 'react';
import { 
  Send, Sparkles, Settings, Key, RefreshCw, Play, 
  CheckCircle2, XCircle, AlertCircle, Database,
  Image, Cpu, Lock, Layout, Terminal, Code2, X
} from 'lucide-react';
import { 
  getWorkbookData, 
  writeFormulas, 
  writeValues, 
  createExcelTable, 
  sortTableColumn, 
  filterTableColumn, 
  createExcelChart, 
  createExcelPivotTable,
  WorkbookData
} from './utils/excelService';
import { getGeminiExcelAction, getDeepSeekExcelAction, ActionPlan } from './utils/geminiService';

interface AppProps {
  isExcel: boolean;
}

interface Message {
  id: string;
  sender: 'ai' | 'user';
  text: string;
  plan?: ActionPlan;
  status?: 'idle' | 'running' | 'completed' | 'failed';
  stepStatuses?: ('pending' | 'running' | 'success' | 'failed')[];
  errorMsg?: string;
}

export default function App({ isExcel }: AppProps) {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState('');
  const [provider, setProvider] = useState<'gemini' | 'deepseek'>(() => {
    const saved = localStorage.getItem('excel_ai_provider');
    return (saved === 'deepseek' || saved === 'gemini') ? saved : 'gemini';
  });
  const [geminiApiKey, setGeminiApiKey] = useState(() => {
    return localStorage.getItem('gemini_excel_api_key') || import.meta.env.VITE_GEMINI_API_KEY || '';
  });
  const [deepseekApiKey, setDeepseekApiKey] = useState(() => {
    return localStorage.getItem('deepseek_excel_api_key') || import.meta.env.VITE_DEEPSEEK_API_KEY || '';
  });
  const [showSettings, setShowSettings] = useState(false);
  const [workbookData, setWorkbookData] = useState<WorkbookData | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncingSheet, setIsSyncingSheet] = useState(false);
  const [showShowcase, setShowShowcase] = useState(!isExcel);
  
  const activeSheetData = workbookData
    ? workbookData.sheets.find(s => s.name === workbookData.activeSheetName) || workbookData.sheets[0]
    : null;
  
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Load API Key dari localStorage saat startup
  useEffect(() => {
    // Tulis defaults ke localStorage jika belum diset
    if (!localStorage.getItem('gemini_excel_api_key') && import.meta.env.VITE_GEMINI_API_KEY) {
      localStorage.setItem('gemini_excel_api_key', import.meta.env.VITE_GEMINI_API_KEY);
    }
    if (!localStorage.getItem('deepseek_excel_api_key') && import.meta.env.VITE_DEEPSEEK_API_KEY) {
      localStorage.setItem('deepseek_excel_api_key', import.meta.env.VITE_DEEPSEEK_API_KEY);
    }
    if (!localStorage.getItem('excel_ai_provider')) {
      localStorage.setItem('excel_ai_provider', 'gemini');
    }
    
    // Tampilkan pesan selamat datang
    setMessages([
      {
        id: 'welcome',
        sender: 'ai',
        text: 'Halo! Saya adalah **AI Excel Assistant** lokal Anda. Saya dapat membantu mengolah data, menulis rumus, menyaring (filter), mengurutkan, membuat pivot table, serta diagram grafik secara otomatis.\n\nSilakan muat data sheet terlebih dahulu dan ketik perintah yang Anda inginkan!'
      }
    ]);

    // Baca data sheet awal
    syncSheetInfo();
  }, []);

  // Scroll chat otomatis ke bawah
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isLoading]);

  // Eksekusi Rencana Aksi secara otomatis (Auto-Inject) saat AI memberikan respon
  useEffect(() => {
    const idleMessage = messages.find(
      m => m.sender === 'ai' && m.status === 'idle' && m.plan && m.plan.actions.length > 0
    );
    if (idleMessage) {
      executePlan(idleMessage.id);
    }
  }, [messages]);

  const syncSheetInfo = async () => {
    if (!isExcel) {
      // Simulasi sheet kosong jika tidak di Excel
      setWorkbookData({
        activeSheetName: "Sheet1",
        sheets: [
          {
            name: "Sheet1",
            hasData: false,
            headers: [],
            rowCount: 0,
            columnCount: 0,
            values: [],
            formulas: [],
            address: "A1",
            columnIndex: 0,
            rowIndex: 0
          }
        ]
      });
      return;
    }
    
    setIsSyncingSheet(true);
    try {
      const data = await getWorkbookData();
      setWorkbookData(data);
    } catch (e: any) {
      console.error(e);
      alert("Gagal memuat data dari Excel: " + (e.message || String(e)));
    } finally {
      setIsSyncingSheet(false);
    }
  };

  const handleSaveApiKey = () => {
    localStorage.setItem('excel_ai_provider', provider);
    localStorage.setItem('gemini_excel_api_key', geminiApiKey.trim());
    localStorage.setItem('deepseek_excel_api_key', deepseekApiKey.trim());
    setShowSettings(false);
    
    const activeKey = provider === 'gemini' ? geminiApiKey : deepseekApiKey;
    const providerName = provider === 'gemini' ? 'Gemini AI' : 'DeepSeek AI';

    setMessages(prev => [
      ...prev,
      {
        id: `sys-${Date.now()}`,
        sender: 'ai',
        text: activeKey.trim() 
          ? `🔑 **Konfigurasi disimpan!** Sekarang saya menggunakan **${providerName}** untuk memproses perintah Excel Anda secara penuh.`
          : `⚠️ **API Key untuk ${providerName} kosong.** Aplikasi sekarang berjalan dalam **Demo Mode** dengan perintah terbatas.`
      }
    ]);
  };

  const handleSend = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!input.trim()) return;

    const userPrompt = input.trim();
    setInput('');
    setIsLoading(true);

    const userMsgId = `user-${Date.now()}`;
    const aiMsgId = `ai-${Date.now()}`;

    // 1. Tambah pesan user ke chat
    setMessages(prev => [
      ...prev,
      { id: userMsgId, sender: 'user', text: userPrompt }
    ]);

    // Ambil data workbook ter-update sebelum kirim ke AI
    let currentWorkbook = workbookData;
    if (isExcel) {
      try {
        currentWorkbook = await getWorkbookData();
        setWorkbookData(currentWorkbook);
      } catch (err) {
        console.error("Gagal sinkronisasi data workbook sebelum mengirim perintah:", err);
      }
    }

    if (!currentWorkbook) {
      setMessages(prev => [
        ...prev,
        {
          id: aiMsgId,
          sender: 'ai',
          text: '❌ Gagal mendeteksi lembar kerja aktif. Mohon pastikan Excel Anda terbuka dan aktif.'
        }
      ]);
      setIsLoading(false);
      return;
    }

    try {
      // 2. Panggil AI Service sesuai provider
      const activeApiKey = provider === 'gemini' ? geminiApiKey.trim() : deepseekApiKey.trim();
      let plan;
      if (provider === 'gemini') {
        plan = await getGeminiExcelAction(userPrompt, currentWorkbook, activeApiKey);
      } else {
        plan = await getDeepSeekExcelAction(userPrompt, currentWorkbook, activeApiKey);
      }
      
      // Buat status langkah default (pending)
      const stepStatuses = plan.actions.map(() => 'pending' as const);

      setMessages(prev => [
        ...prev,
        {
          id: aiMsgId,
          sender: 'ai',
          text: plan.explanation,
          plan: plan,
          status: plan.actions.length > 0 ? 'idle' : 'completed',
          stepStatuses: stepStatuses
        }
      ]);
    } catch (err: any) {
      setMessages(prev => [
        ...prev,
        {
          id: aiMsgId,
          sender: 'ai',
          text: `❌ Terjadi kesalahan saat memproses perintah Anda:\n\`\`\`\n${err.message || err}\n\`\`\``
        }
      ]);
    } finally {
      setIsLoading(false);
    }
  };

  // Eksekusi Action Plan secara berurutan di Microsoft Excel
  const executePlan = async (messageId: string) => {
    const msgIndex = messages.findIndex(m => m.id === messageId);
    if (msgIndex === -1 || !messages[msgIndex].plan) return;

    const message = messages[msgIndex];
    const plan = message.plan!;
    
    // Perbarui status chat menjadi running
    setMessages(prev => {
      const copy = [...prev];
      copy[msgIndex] = {
        ...message,
        status: 'running',
        stepStatuses: plan.actions.map(() => 'pending')
      };
      return copy;
    });

    const stepStatuses = [...(message.stepStatuses || [])];

    try {
      for (let i = 0; i < plan.actions.length; i++) {
        // Set status step saat ini ke running
        stepStatuses[i] = 'running';
        setMessages(prev => {
          const copy = [...prev];
          copy[msgIndex] = { ...copy[msgIndex], stepStatuses: [...stepStatuses] };
          return copy;
        });

        const action = plan.actions[i];
        
        if (!isExcel) {
          // Jika tidak dijalankan dalam Excel (Simulasi/Browser Dev)
          await new Promise(r => setTimeout(r, 1000));
        } else {
          // Jalankan API Excel sesungguhnya
          switch (action.type) {
            case 'CREATE_TABLE':
              await createExcelTable(action.payload.range, action.payload.name);
              break;
            case 'WRITE_FORMULAS':
              await writeFormulas(action.payload.range, action.payload.formulas);
              break;
            case 'WRITE_VALUES':
              await writeValues(action.payload.range, action.payload.values);
              break;
            case 'SORT':
              await sortTableColumn(action.payload.tableName, action.payload.columnName, action.payload.direction);
              break;
            case 'FILTER':
              await filterTableColumn(action.payload.tableName, action.payload.columnName, action.payload.value, action.payload.operator);
              break;
            case 'CREATE_CHART':
              await createExcelChart(action.payload.type, action.payload.range, action.payload.title);
              break;
            case 'CREATE_PIVOT':
              await createExcelPivotTable(action.payload.source, action.payload.rowFields, action.payload.dataFields);
              break;
            default:
              throw new Error(`Aksi tidak dikenal: ${action.type}`);
          }
        }

        // Set status step saat ini ke success
        stepStatuses[i] = 'success';
        setMessages(prev => {
          const copy = [...prev];
          copy[msgIndex] = { ...copy[msgIndex], stepStatuses: [...stepStatuses] };
          return copy;
        });
      }

      // Selesai sukses
      setMessages(prev => {
        const copy = [...prev];
        copy[msgIndex] = {
          ...copy[msgIndex],
          status: 'completed'
        };
        return copy;
      });

      // Singkronisasi ulang data sheet untuk memperbarui info UI
      await syncSheetInfo();

    } catch (error: any) {
      console.error(error);
      // Gagal
      const failedIndex = stepStatuses.findIndex(s => s === 'running');
      if (failedIndex !== -1) {
        stepStatuses[failedIndex] = 'failed';
      }
      setMessages(prev => {
        const copy = [...prev];
        copy[msgIndex] = {
          ...copy[msgIndex],
          status: 'failed',
          stepStatuses: [...stepStatuses],
          errorMsg: error.message || String(error)
        };
        return copy;
      });
    }
  };

  const loadSampleData = async () => {
    if (!isExcel) {
      alert("Operasi ini hanya berfungsi di dalam aplikasi Microsoft Excel Desktop.");
      return;
    }
    
    setIsLoading(true);
    try {
      // Tulis header & data sampel
      await writeValues("A1:E7", [
        ["Tanggal Pemesanan", "No. Pemesanan", "Nama Customer", "Domisili", "Lama Pengiriman"],
        ["6-Mar-2021", 106, "Budi", "Jakarta", 3],
        ["2-Mar-2021", 102, "Siti", "Bandung", 2],
        ["4-Mar-2021", 104, "Andi", "Jakarta", 5],
        ["5-Mar-2021", 105, "Dewi", "Surabaya", 1],
        ["10-Mar-2021", 110, "Rian", "Bandung", 4],
        ["3-Mar-2021", 103, "Lina", "Jakarta", 2]
      ]);
      await syncSheetInfo();
      
      setMessages(prev => [
        ...prev,
        {
          id: `sample-${Date.now()}`,
          sender: 'ai',
          text: '📝 **Data sampel berhasil dimasukkan!**\n\nSekarang Anda bisa mencoba perintah seperti:\n- *Urutkan kolom Lama Pengiriman dari terkecil*\n- *Filter domisili Jakarta*\n- *Buat tabel pivot*'
        }
      ]);
    } catch (e: any) {
      alert("Gagal memuat data sampel: " + e.message);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header">
        <div className="brand-section">
          <img src="/icon.png" alt="Excel AI Logo" className="app-logo" onError={(e) => { e.currentTarget.src = 'https://img.icons8.com/color/96/microsoft-excel-2019.png' }} />
          <div className="app-title">
            <h1>AI Assistant</h1>
            <span>Excel Automation</span>
          </div>
        </div>
        
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          <button 
            className={`suggestion-chip ${!isExcel && !showShowcase ? 'pulse' : ''}`}
            style={{ 
              display: 'flex', 
              alignItems: 'center', 
              gap: '4px', 
              background: showShowcase ? 'rgba(139, 92, 246, 0.15)' : 'rgba(255, 255, 255, 0.04)', 
              borderColor: showShowcase ? 'rgba(139, 92, 246, 0.3)' : 'var(--border-color)',
              color: showShowcase ? '#c084fc' : 'var(--text-secondary)'
            }}
            onClick={() => setShowShowcase(!showShowcase)}
            title="Tampilkan Portofolio & Fitur Proyek"
          >
            <Sparkles size={12} />
            Info & Fitur
          </button>

          <button 
            className="suggestion-chip" 
            style={{ display: 'flex', alignItems: 'center', gap: '4px', background: 'rgba(16, 185, 129, 0.1)', borderColor: 'rgba(16, 185, 129, 0.3)' }}
            onClick={syncSheetInfo}
            title="Sinkronisasi Data Excel"
          >
            <RefreshCw size={12} className={isSyncingSheet ? 'spin' : ''} />
            Muat Sheet
          </button>
          
          <div className="status-badge">
            <span className={`status-dot ${isExcel ? 'active' : 'offline'}`}></span>
            {isExcel ? 'Terhubung ke Excel' : 'Browser Mode (Demo)'}
          </div>
          
          <button 
            onClick={() => setShowSettings(!showSettings)} 
            className="btn-send"
            style={{ width: '32px', height: '32px', background: 'rgba(255,255,255,0.06)' }}
          >
            <Settings size={14} />
          </button>
        </div>
      </header>

      {/* Main Layout Area */}
      <div className="main-content-layout">
        {/* Main Chat Pane */}
        <div className="chat-pane" style={{ flex: 1, height: '100%' }}>
        
        {/* Panel API Key Settings */}
        {showSettings && (
          <div className="settings-bar" style={{ padding: '14px 18px', background: 'rgba(15, 23, 42, 0.8)', borderBottom: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <div className="settings-toggle">
              <span style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>
                <Key size={14} style={{ color: 'var(--accent-secondary)' }} />
                AI Provider & API Key Configuration
              </span>
            </div>
            
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '10px', alignItems: 'center' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', flex: '1 1 120px' }}>
                <label style={{ fontSize: '10px', color: 'var(--text-secondary)', fontWeight: 600, textTransform: 'uppercase' }}>AI Provider</label>
                <select 
                  value={provider} 
                  onChange={(e) => setProvider(e.target.value as 'gemini' | 'deepseek')}
                  style={{
                    background: 'rgba(255, 255, 255, 0.05)',
                    border: '1px solid var(--border-color)',
                    padding: '8px 12px',
                    borderRadius: '6px',
                    color: 'var(--text-primary)',
                    fontSize: '12px',
                    outline: 'none',
                    cursor: 'pointer'
                  }}
                >
                  <option value="gemini" style={{ background: '#0f172a' }}>Google Gemini</option>
                  <option value="deepseek" style={{ background: '#0f172a' }}>DeepSeek AI</option>
                </select>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', flex: '2 1 200px' }}>
                <label style={{ fontSize: '10px', color: 'var(--text-secondary)', fontWeight: 600, textTransform: 'uppercase' }}>
                  {provider === 'gemini' ? 'Gemini API Key' : 'DeepSeek API Key'}
                </label>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <input 
                    type="password" 
                    className="api-input" 
                    placeholder={provider === 'gemini' ? "Masukkan API Key Gemini..." : "Masukkan API Key DeepSeek..."}
                    value={provider === 'gemini' ? geminiApiKey : deepseekApiKey}
                    onChange={(e) => {
                      if (provider === 'gemini') {
                        setGeminiApiKey(e.target.value);
                      } else {
                        setDeepseekApiKey(e.target.value);
                      }
                    }}
                    style={{ flex: 1 }}
                  />
                  <button onClick={handleSaveApiKey} className="btn-save">
                    Simpan
                  </button>
                </div>
              </div>
            </div>

            <div style={{ fontSize: '10px', color: 'var(--text-muted)' }}>
              {provider === 'gemini' 
                ? 'Model: gemini-2.5-flash. Dapatkan API Key gratis di Google AI Studio.' 
                : 'Model: deepseek-chat. Menggunakan API key default dari pengguna.'}
            </div>
          </div>
        )}

        {/* Notifikasi Sheet Kosong */}
        {activeSheetData && !activeSheetData.hasData && isExcel && (
          <div style={{ padding: '12px 16px 0 16px' }}>
            <div className="info-banner">
              <Database size={14} style={{ color: '#10b981' }} />
              <div>
                Lembar kerja Excel Anda terdeteksi kosong. 
                <button 
                  onClick={loadSampleData} 
                  style={{ 
                    background: 'none', 
                    border: 'none', 
                    color: '#10b981', 
                    textDecoration: 'underline', 
                    fontWeight: 700, 
                    cursor: 'pointer', 
                    marginLeft: '6px' 
                  }}
                >
                  Isikan Data Sampel
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Messages */}
        <div className="messages-container">
          {messages.map((msg) => (
            <div key={msg.id} className={`message-bubble ${msg.sender}`}>
              <div className="message-sender">
                {msg.sender === 'ai' ? 'AI Assistant' : 'Anda'}
              </div>
              
              <div 
                style={{ whiteSpace: 'pre-wrap' }} 
                dangerouslySetInnerHTML={{ __html: formatMarkdown(msg.text) }}
              />

              {/* Render Action Plan jika ada */}
              {msg.plan && msg.plan.actions.length > 0 && (
                <div className="action-plan-container">
                  <div className="action-plan-title">Rencana Aksi AI:</div>
                  
                  {msg.plan.actions.map((act, i) => (
                    <div key={i} className="action-step">
                      <div className="step-status-icon">
                        {msg.stepStatuses && msg.stepStatuses[i] === 'pending' && (
                          <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: '#64748b' }}></div>
                        )}
                        {msg.stepStatuses && msg.stepStatuses[i] === 'running' && (
                          <RefreshCw size={12} className="spin" style={{ color: '#8b5cf6' }} />
                        )}
                        {msg.stepStatuses && msg.stepStatuses[i] === 'success' && (
                          <CheckCircle2 size={12} style={{ color: '#10b981' }} />
                        )}
                        {msg.stepStatuses && msg.stepStatuses[i] === 'failed' && (
                          <XCircle size={12} style={{ color: '#f43f5e' }} />
                        )}
                      </div>
                      <div className="step-text" style={{ 
                        textDecoration: msg.stepStatuses && msg.stepStatuses[i] === 'success' ? 'line-through' : 'none',
                        color: msg.stepStatuses && msg.stepStatuses[i] === 'success' ? '#64748b' : '#f8fafc'
                      }}>
                        {getStepDescription(act)}
                      </div>
                    </div>
                  ))}

                  {msg.status === 'idle' && (
                    <button onClick={() => executePlan(msg.id)} className="btn-execute">
                      <Play size={12} />
                      Jalankan Aksi di Excel
                    </button>
                  )}

                  {msg.status === 'running' && (
                    <div className="btn-execute" style={{ background: '#3b0764', cursor: 'default' }}>
                      <RefreshCw size={12} className="spin" />
                      Sedang memproses...
                    </div>
                  )}

                  {msg.status === 'completed' && (
                    <div style={{ marginTop: '8px', display: 'flex', alignItems: 'center', gap: '6px', color: '#10b981', fontSize: '11px', fontWeight: 600 }}>
                      <CheckCircle2 size={12} />
                      Berhasil dijalankan di Excel!
                    </div>
                  )}

                  {msg.status === 'failed' && (
                    <div style={{ marginTop: '8px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: '#f43f5e', fontSize: '11px', fontWeight: 600 }}>
                        <AlertCircle size={12} />
                        Gagal memproses langkah.
                      </div>
                      {msg.errorMsg && (
                        <div style={{ fontSize: '10px', color: '#fda4af', marginTop: '4px', background: 'rgba(244,63,94,0.1)', padding: '6px', borderRadius: '4px' }}>
                          Detail: {msg.errorMsg}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
          ))}
          
          {isLoading && (
            <div className="message-bubble ai">
              <div className="message-sender">AI Assistant</div>
              <div className="typing-indicator">
                <div className="typing-dot"></div>
                <div className="typing-dot"></div>
                <div className="typing-dot"></div>
              </div>
            </div>
          )}
          
          <div ref={messagesEndRef} />
        </div>

        {/* Input Area */}
        <div className="chat-input-area">


          <form onSubmit={handleSend} className="chat-form">
            <input 
              type="text" 
              className="chat-input" 
              placeholder={((provider === 'gemini' ? geminiApiKey : deepseekApiKey).trim()) ? "Ketik perintah untuk Excel..." : "Tulis perintah (Demo Mode)..."}
              value={input}
              onChange={(e) => setInput(e.target.value)}
              disabled={isLoading}
            />
            <button type="submit" className="btn-send" disabled={isLoading || !input.trim()}>
              <Send size={16} />
            </button>
          </form>
          
          {!((provider === 'gemini' ? geminiApiKey : deepseekApiKey).trim()) && (
            <div style={{ fontSize: '10px', color: '#8b5cf6', display: 'flex', alignItems: 'center', gap: '4px', justifyContent: 'center' }}>
              <Sparkles size={10} />
              Berjalan dalam Demo Mode. Buka Pengaturan ⚙️ untuk memasukkan API Key {provider === 'gemini' ? 'Gemini' : 'DeepSeek'}.
            </div>
          )}
        </div>
        </div>

        {/* Recruiter Showcase Panel */}
        {showShowcase && (
          <RecruiterShowcase onClose={() => setShowShowcase(false)} />
        )}
      </div>
    </div>
  );
}



// Parser markdown sederhana untuk chat text bubble
function formatMarkdown(text: string): string {
  let html = text;
  // Bold **text**
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  // Italic *text*
  html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
  // Inline code `code`
  html = html.replace(/`(.*?)`/g, '<code style="background:rgba(255,255,255,0.06); padding:2px 4px; border-radius:4px; font-family:monospace;">$1</code>');
  // List items
  html = html.replace(/^- (.*?)$/gm, '<li style="margin-left: 15px; margin-top: 4px;">$1</li>');
  return html;
}

// Dapatkan deskripsi langkah yang ramah untuk dibaca pengguna di UI
function getStepDescription(action: any): string {
  switch (action.type) {
    case 'CREATE_TABLE':
      return `Format area data ${action.payload.range} menjadi Excel Table "${action.payload.name}"`;
    case 'WRITE_FORMULAS':
      return `Tulis rumus Excel pada kolom ${action.payload.range}`;
    case 'WRITE_VALUES':
      return `Tulis nilai data pada sel ${action.payload.range}`;
    case 'SORT':
      return `Urutkan tabel ${action.payload.tableName} berdasarkan kolom ${action.payload.columnName} (${action.payload.direction === 'asc' ? 'Terkecil' : 'Terbesar'})`;
    case 'FILTER':
      if (action.payload.operator === 'Clear') {
        return `Bersihkan filter kolom ${action.payload.columnName} pada tabel ${action.payload.tableName}`;
      }
      return `Filter tabel ${action.payload.tableName} kolom ${action.payload.columnName} = "${action.payload.value}"`;
    case 'CREATE_CHART':
      return `Buat grafik ${action.payload.type} dari data ${action.payload.range}`;
    case 'CREATE_PIVOT':
      return `Buat Pivot Table baru dari sumber data ${action.payload.source}`;
    default:
      return `Eksekusi aksi: ${action.type}`;
  }
}

interface RecruiterShowcaseProps {
  onClose: () => void;
}

function RecruiterShowcase({ onClose }: RecruiterShowcaseProps) {
  const [activeTab, setActiveTab] = React.useState<'screenshot' | 'features' | 'tech'>('screenshot');

  return (
    <div className="showcase-panel">
      <div className="showcase-header">
        <h2>
          <Sparkles size={15} style={{ color: 'var(--accent-secondary)' }} />
          Portofolio & Fitur Proyek
        </h2>
        <button onClick={onClose} className="showcase-close-btn" title="Tutup Panel">
          <X size={15} />
        </button>
      </div>

      <div className="showcase-tabs">
        <button 
          className={`showcase-tab-btn ${activeTab === 'screenshot' ? 'active' : ''}`}
          onClick={() => setActiveTab('screenshot')}
        >
          <Image size={13} />
          Demo & Mockup
        </button>
        <button 
          className={`showcase-tab-btn ${activeTab === 'features' ? 'active' : ''}`}
          onClick={() => setActiveTab('features')}
        >
          <Layout size={13} />
          Fitur Utama
        </button>
        <button 
          className={`showcase-tab-btn ${activeTab === 'tech' ? 'active' : ''}`}
          onClick={() => setActiveTab('tech')}
        >
          <Cpu size={13} />
          Teknologi & Arsitektur
        </button>
      </div>

      <div className="showcase-content">
        {activeTab === 'screenshot' && (
          <>
            <div className="showcase-section-title">
              <Image size={12} />
              Tampilan Add-in di Excel
            </div>
            
            <div className="screenshot-card">
              <div className="screenshot-image-wrapper">
                <img 
                  src="/excel_addin_showcase.png" 
                  alt="Excel AI Assistant Showcase" 
                  className="screenshot-image"
                  onError={(e) => {
                    e.currentTarget.src = 'https://placehold.co/600x400/0f172a/f8fafc?text=AI+Excel+Assistant+Mockup';
                  }}
                />
                <span className="screenshot-badge">Simulasi Excel</span>
              </div>
              <div className="screenshot-info">
                <h3>Integrasi Side-by-Side dengan Excel</h3>
                <p>
                  Gambar di atas menyimulasikan bagaimana AI Excel Assistant berjalan sebagai panel samping (taskpane) di dalam Microsoft Excel Desktop maupun Excel Web. Pengguna dapat mengobrol dengan AI, dan AI akan mengontrol spreadsheet secara real-time dan aman.
                </p>
              </div>
            </div>

            <div className="architecture-card" style={{ marginTop: '10px' }}>
              <Terminal size={20} style={{ color: 'var(--accent-primary)', flexShrink: 0 }} />
              <div className="architecture-text">
                <h4>Siap untuk HR & Recruiter Review</h4>
                <p style={{ fontSize: '11px', lineHeight: '1.5' }}>
                  Aplikasi ini dirancang dengan standar kualitas tinggi: kode terstruktur menggunakan TypeScript, integrasi API AI yang modular, dan kepatuhan penuh terhadap Office Add-in manifest.
                </p>
              </div>
            </div>
          </>
        )}

        {activeTab === 'features' && (
          <>
            <div className="showcase-section-title">
              <Layout size={12} />
              Kemampuan Otomatisasi AI
            </div>

            <div className="features-grid">
              <div className="feature-card">
                <div className="feature-icon-wrapper">
                  <Code2 size={15} />
                </div>
                <div className="feature-details">
                  <h3>1. Generator Rumus Instan</h3>
                  <p>Menghasilkan rumus Excel yang kompleks (seperti SUMIFS, VLOOKUP, INDEX/MATCH) berdasarkan perintah bahasa sehari-hari dan langsung menulisnya ke sel.</p>
                  <span className="feature-api-badge">range.formulas = [...]</span>
                </div>
              </div>

              <div className="feature-card">
                <div className="feature-icon-wrapper" style={{ color: 'var(--accent-secondary)', borderColor: 'rgba(139, 92, 246, 0.2)', background: 'rgba(139, 92, 246, 0.08)' }}>
                  <Database size={15} />
                </div>
                <div className="feature-details">
                  <h3>2. Pembuat Excel Table Otomatis</h3>
                  <p>Mendeteksi jangkauan data (range) secara otomatis dan memformat data mentah menjadi Excel Table resmi dengan sekali klik.</p>
                  <span className="feature-api-badge">worksheet.tables.add()</span>
                </div>
              </div>

              <div className="feature-card">
                <div className="feature-icon-wrapper">
                  <RefreshCw size={15} />
                </div>
                <div className="feature-details">
                  <h3>3. Pengurutan & Penyaringan Data</h3>
                  <p>Mendukung pengurutan kolom secara naik/turun dan menerapkan filter kustom pada tabel Excel secara terprogram melalui perintah AI.</p>
                  <span className="feature-api-badge">table.sort / table.column.filter</span>
                </div>
              </div>

              <div className="feature-card">
                <div className="feature-icon-wrapper" style={{ color: 'var(--accent-secondary)', borderColor: 'rgba(139, 92, 246, 0.2)', background: 'rgba(139, 92, 246, 0.08)' }}>
                  <Sparkles size={15} />
                </div>
                <div className="feature-details">
                  <h3>4. Pembuat Grafik & Diagram</h3>
                  <p>Membaca data sheet lalu membuat grafik representatif (batang, kolom, garis, atau lingkaran) secara otomatis untuk visualisasi cepat.</p>
                  <span className="feature-api-badge">worksheet.charts.add()</span>
                </div>
              </div>

              <div className="feature-card">
                <div className="feature-icon-wrapper">
                  <Layout size={15} />
                </div>
                <div className="feature-details">
                  <h3>5. Generator Pivot Table</h3>
                  <p>Membuat ringkasan Pivot Table interaktif dari database yang besar dengan konfigurasi field baris dan kolom yang disesuaikan.</p>
                  <span className="feature-api-badge">worksheet.pivotTables.add()</span>
                </div>
              </div>

              <div className="feature-card">
                <div className="feature-icon-wrapper" style={{ color: 'var(--accent-secondary)', borderColor: 'rgba(139, 92, 246, 0.2)', background: 'rgba(139, 92, 246, 0.08)' }}>
                  <Cpu size={15} />
                </div>
                <div className="feature-details">
                  <h3>6. Sinkronisasi Data Sheet Dinamis</h3>
                  <p>Membaca metadata sheet aktif, header kolom, jumlah baris, dan rumus yang terpasang agar AI memahami konteks data sebelum memberikan aksi.</p>
                  <span className="feature-api-badge">worksheet.getSelectedRange()</span>
                </div>
              </div>
            </div>
          </>
        )}

        {activeTab === 'tech' && (
          <>
            <div className="showcase-section-title">
              <Cpu size={12} />
              Teknologi & Keunggulan Kode
            </div>

            <div className="tech-container">
              <div className="tech-group">
                <div className="tech-group-title" style={{ color: 'var(--accent-primary)' }}>
                  <Layout size={12} style={{ color: 'var(--accent-primary)' }} />
                  Frontend & Integrasi
                </div>
                <div className="tech-badge-container">
                  <span className="tech-badge">React 18</span>
                  <span className="tech-badge">TypeScript</span>
                  <span className="tech-badge">Vite Builder</span>
                  <span className="tech-badge">Office JS SDK</span>
                  <span className="tech-badge">Lucide Icons</span>
                  <span className="tech-badge">Custom CSS Variables</span>
                </div>
              </div>

              <div className="tech-group">
                <div className="tech-group-title" style={{ color: 'var(--accent-secondary)' }}>
                  <Cpu size={12} style={{ color: 'var(--accent-secondary)' }} />
                  AI Orchestrator Engine
                </div>
                <div className="tech-badge-container">
                  <span className="tech-badge">Google Gemini API</span>
                  <span className="tech-badge">DeepSeek AI API</span>
                  <span className="tech-badge">Structured JSON Output</span>
                  <span className="tech-badge">Modular Command Pattern</span>
                </div>
              </div>

              <div className="tech-group">
                <div className="tech-group-title" style={{ color: '#f43f5e' }}>
                  <Lock size={12} style={{ color: '#f43f5e' }} />
                  Keamanan & Praktik Terbaik
                </div>
                <div className="tech-badge-container">
                  <span className="tech-badge">Local API Keys</span>
                  <span className="tech-badge">No Third-party DB</span>
                  <span className="tech-badge">Transaction Batching</span>
                  <span className="tech-badge">Graceful Fallbacks</span>
                </div>
              </div>

              <div className="architecture-card">
                <Lock size={20} style={{ color: 'var(--accent-primary)', flexShrink: 0 }} />
                <div className="architecture-text">
                  <h4>Privasi Data Terjamin</h4>
                  <p style={{ fontSize: '11px', lineHeight: '1.5' }}>
                    Kunci API disimpan secara lokal di browser (`localStorage`) dan tidak pernah dikirim ke server pihak ketiga mana pun. Kunci ini hanya digunakan untuk mengautentikasi langsung ke Google / DeepSeek API.
                  </p>
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
