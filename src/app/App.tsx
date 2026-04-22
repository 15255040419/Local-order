// Import polyfills first for browser compatibility
import '../polyfills';
import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import {
  Upload, FileCheck, Download, Trash2, CheckCircle2, Loader2,
  Sun, Moon, ChevronDown, ChevronUp, X, AlertTriangle, Terminal
} from 'lucide-react';
import { Toaster, toast } from 'sonner';

/* ─────────────────────────── Constants ─────────────────────────── */
const DB_NAME = 'FTH_Orders_DB_V3';
const STORE_NAME = 'files';

/* ─────────────────────────── Types ─────────────────────────── */
interface FileCache { name: string; data: ArrayBuffer; }
interface StoreInfo {
  storeName: string; storeCode: string; address: string;
  contact: string; phone: string; _isDup?: boolean;
}
interface LogEntry { time: string; message: string; type: 'info' | 'success' | 'warn' | 'error'; }
interface PreviewData { title: string; headers: string[]; rows: any[]; stats: string; }
interface ErrorModal { title: string; desc?: string; }

const getDateStr = () => new Date().toISOString().split('T')[0].replace(/-/g, '');

/* ─────────────────────────── Theme helpers ─────────────────────────── */
const T = {
  bg: (d: boolean) => d ? 'bg-[#1c1c1e]' : 'bg-[#f2f2f7]',
  card: (d: boolean) => d ? 'bg-[#2c2c2e]' : 'bg-white',
  cardHover: (d: boolean) => d ? 'hover:bg-[#3a3a3c]' : 'hover:bg-gray-50',
  border: (d: boolean) => d ? 'border-white/10' : 'border-black/8',
  text: (d: boolean) => d ? 'text-white' : 'text-gray-900',
  textSub: (d: boolean) => d ? 'text-gray-400' : 'text-gray-500',
  textMuted: (d: boolean) => d ? 'text-gray-500' : 'text-gray-400',
  input: (d: boolean) => d ? 'bg-[#3a3a3c] border-white/10 text-white' : 'bg-gray-100 border-black/8 text-gray-900',
  logBg: (d: boolean) => d ? 'bg-[#1a1a1a]' : 'bg-gray-900',
};

/* ─────────────────────────── Province Helper ─────────────────────────── */
const PROVINCE_LIST = [
  '北京', '天津', '上海', '重庆', '河北', '山西', '辽宁', '吉林', '黑龙江', '江苏', '浙江', '安徽',
  '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '海南', '四川', '贵州', '云南', '陕西',
  '甘肃', '青海', '台湾', '内蒙古', '广西', '西藏', '宁夏', '新疆', '香港', '澳门'
];

const extractProvince = (address: string) => {
  if (!address) return null;
  const cleanAddr = String(address).replace(/\s+/g, '');
  for (const p of PROVINCE_LIST) {
    if (cleanAddr.includes(p)) return p;
  }
  return null;
};

const completeAddress = (addr: any, prov: any, city: any, dist: any) => {
  const a = String(addr || '').trim();
  const p = String(prov || '').trim();
  const c = String(city || '').trim();
  const d = String(dist || '').trim();

  if (!a || a === 'nan') {
    const parts = [];
    if (p && !['nan', '省直辖县级行政区划'].includes(p)) parts.push(p);
    if (c && !['nan', '市辖区', '省直辖县级行政区划'].includes(c)) parts.push(c);
    if (d && d !== 'nan') parts.push(d);
    return parts.join('');
  }

  const pKey = p.replace(/(省|市|自治区|自治州|特别行政区)$/, '');
  const hasP = pKey && (a.startsWith(p) || a.startsWith(pKey));
  const cKey = c.replace(/市$/, '');
  const hasC = cKey && cKey !== '市辖区' && a.slice(0, 15).includes(cKey);
  const hasD = d && d !== '市辖区' && a.includes(d.replace(/(区|县|市|县)$/, ''));

  if (hasP && hasC && (hasD || d === '市辖区')) return a;

  let result = a;
  if (hasP && hasC && !hasD && d && d !== '市辖区') {
    const cIdx = a.indexOf(cKey);
    if (cIdx >= 0) {
      let insPos = cIdx + cKey.length;
      if (a[insPos] === '市') insPos++;
      result = a.slice(0, insPos) + d + a.slice(insPos);
    } else { result = a + d; }
  } else if (hasC && !hasP && p) {
    result = p + a;
  } else if (hasP && !hasC && pKey) {
    const afterP = a.startsWith(p) ? a.slice(p.length) : a.slice(pKey.length);
    let ins = (c && c !== '市辖区') ? c : '';
    if (!hasD && d && d !== '市辖区') ins += d;
    result = p + ins + afterP;
  } else {
    let ins = '';
    if (!hasP && p) ins += p;
    if (!hasC && c && c !== '市辖区') ins += c;
    if (!hasD && d && d !== '市辖区') ins += d;
    result = ins + a;
  }
  return result;
};

/* ═══════════════════════════ App ═══════════════════════════ */
export default function App() {
  const [isDark, setIsDark] = useState(false);
  const [activeTab, setActiveTab] = useState<'fulujia' | 'shangsong'>('fulujia');
  const [currentStep, setCurrentStep] = useState(1);
  const [expandedStep, setExpandedStep] = useState<number | null>(1);
  const [expandedSS, setExpandedSS] = useState(true);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [showLog, setShowLog] = useState(false);
  const [previewData, setPreviewData] = useState<PreviewData | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<Record<string, File | null>>({});
  const [cachedFiles, setCachedFiles] = useState<Record<string, boolean>>({});
  const [errorModal, setErrorModal] = useState<ErrorModal | null>(null);

  const dbRef = useRef<IDBDatabase | null>(null);
  const memExtractedStoresRef = useRef<StoreInfo[] | null>(null);
  const memPurchaseOrderWBRef = useRef<any>(null);
  const currentDownloadFnRef = useRef<(() => void) | null>(null);
  const resultsMemoryRef = useRef<Record<string, { preview: PreviewData; download: () => void; filename?: string }>>({});
  const logContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => { initDB(); }, []);
  useEffect(() => {
    if (logContainerRef.current) {
      logContainerRef.current.scrollTop = logContainerRef.current.scrollHeight;
    }
  }, [logs]);

  /* ── DB ── */
  const initDB = async () => {
    return new Promise<void>((resolve) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = (e) => {
        const db = (e.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
      };
      req.onsuccess = async (e) => {
        dbRef.current = (e.target as IDBOpenDBRequest).result;
        await loadCachedFilesList();
        resolve();
      };
    });
  };

  const loadCachedFilesList = async () => {
    if (!dbRef.current) return;
    const fileIds = ['f-hist', 'f-tpl', 'f-order', 'f-jike', 'f-menu', 's-tpl', 's-price'];
    const cached: Record<string, boolean> = {};
    for (const id of fileIds) { const f = await getFileDB(id); if (f) cached[id] = true; }
    setCachedFiles(cached);
  };

  const saveFileDB = async (id: string, file: File) => {
    if (!dbRef.current) return;
    const buffer = await file.arrayBuffer();
    return new Promise<void>((resolve) => {
      const tx = dbRef.current!.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put({ name: file.name, data: buffer }, id);
      tx.oncomplete = () => resolve();
    });
  };

  const getFileDB = async (id: string): Promise<FileCache | null> => {
    if (!dbRef.current) return null;
    return new Promise((resolve) => {
      const tx = dbRef.current!.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(id);
      req.onsuccess = () => resolve(req.result || null);
    });
  };

  const getWorkbook = async (inputId: string) => {
    const file = uploadedFiles[inputId];
    if (file) { const buffer = await file.arrayBuffer(); return XLSX.read(buffer, { type: 'array' }); }
    const cached = await getFileDB(inputId);
    return cached ? XLSX.read(cached.data, { type: 'array' }) : null;
  };

  /* ── Error Modal ── */
  const showError = (title: string, desc?: string) => setErrorModal({ title, desc });

  /* ── Duplicate Warning Toast ── */
  const showDupWarning = (count: number, context: string) => {
    if (count === 0) return;
    toast.warning(`发现 ${count} 条重复项`, {
      description: `${context} — 已用红色标注，请人工核对后再导出`,
      duration: 6000,
    });
  };

  /* ── Log ── */
  const log = (message: string, type: LogEntry['type'] = 'info') => {
    const time = new Date().toLocaleTimeString('zh-CN');
    setLogs(prev => [...prev, { time, message, type }]);
  };

  const logHtml = (entry: LogEntry): string => {
    const colorMap = {
      success: '#30d158',
      warn: '#ffd60a',
      error: '#ff453a',
      info: '#ebebf5cc',
    };
    const color = colorMap[entry.type];
    const prefix = entry.type === 'success' ? '✓' : entry.type === 'warn' ? '⚠' : entry.type === 'error' ? '✕' : '›';
    return `<span style="color:#636366">[${entry.time}]</span> <span style="color:${color}">${prefix}</span> <span style="color:${color}">${entry.message}</span>`;
  };

  /* ── File Validation ── */
  const validateFile = async (file: File, idKey: string, step: number, _tab: string): Promise<boolean> => {
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array', sheetRows: 20 });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
      const headers = rows[0] ? rows[0].map((h: any) => String(h).trim()) : [];
      const allContentStr = rows.flat().map(c => String(c)).join(' ');

      let reqHeaders: string[] = [];
      let checkFn: (() => boolean) | null = null;
      let fileNameDesc = '';

      if (idKey.startsWith('f-main-s')) {
        const sNum = parseInt(idKey.slice(-1));
        if (sNum === 1) { reqHeaders = ['门店名称', '门店编码', '物料编码']; fileNameDesc = '【第1步主表】 厂家直发采购订单.xlsx'; }
        else if (sNum === 2) { checkFn = () => allContentStr.includes('门店名称') || allContentStr.includes('门店编号'); fileNameDesc = '【第2步主表】 福鹿家提取后门店.xlsx'; }
        else if (sNum === 3) { reqHeaders = ['OMS厂家直发单号', '快递单号']; fileNameDesc = '【第3步主表】 厂家直发采购快递单.xlsx'; }
      } else if (idKey === 'f-hist') {
        checkFn = () => allContentStr.includes('编号') || allContentStr.includes('编码'); fileNameDesc = '【查重文件】 福鹿家总发货记录.xlsx';
      } else if (idKey === 'f-tpl' || idKey === 's-tpl') {
        checkFn = () => wb.SheetNames.includes('货品') || headers.includes('货品名称'); fileNameDesc = '【模板文件】 订单分表模版.xlsx';
      } else if (idKey === 'f-order') {
        reqHeaders = ['OMS厂家直发单号', '门店收货人']; fileNameDesc = '【关联文件】 厂家直发采购订单.xlsx';
      } else if (idKey === 'f-jike') {
        reqHeaders = ['收货人', '物流单号']; fileNameDesc = '【单号源】 快递单号-吉客云.xlsx';
      } else if (idKey === 'f-menu') {
        checkFn = () => /KY|DPK|SF|YT|ZTO|STO|JT/i.test(allContentStr); fileNameDesc = '【单号源】 快递单号-菜单屏.xlsx';
      } else if (idKey === 's-main-input') {
        checkFn = () => allContentStr.includes('订单') || /\d{6}-\d+/.test(allContentStr) || /\d{12,}/.test(allContentStr); fileNameDesc = '【商颂主表】 客户发来的下单记录';
      } else if (idKey === 's-price') {
        checkFn = () => allContentStr.includes('淘宝') || allContentStr.includes('省份'); fileNameDesc = '【规则文件】 快递价格明细表.xls';
      }

      if (reqHeaders.length > 0) {
        const missing = reqHeaders.filter(h => !headers.includes(h));
        if (missing.length > 0) {
          showError(`文件格式不匹配：缺少列 "${missing.join(', ')}"`, `当前需要上传：${fileNameDesc}`);
          return false;
        }
      }
      if (checkFn && !checkFn()) {
        showError('文件内容不匹配', `当前需要上传：${fileNameDesc}`);
        return false;
      }
      return true;
    } catch {
      showError('文件读取失败，可能已损坏或格式不支持');
      return false;
    }
  };

  /* ── File Upload ── */
  const handleFileUpload = async (file: File, idKey: string, shouldPersist = false) => {
    const isValid = await validateFile(file, idKey, currentStep, activeTab);
    if (!isValid) return;
    setUploadedFiles(prev => ({ ...prev, [idKey]: file }));
    log(`成功加载文件: ${file.name}`, 'info');

    if (idKey === 'f-main-s1') {
      const buffer = await file.arrayBuffer();
      memPurchaseOrderWBRef.current = XLSX.read(buffer, { type: 'array' });
      memExtractedStoresRef.current = null;
      delete resultsMemoryRef.current['f1'];
      delete resultsMemoryRef.current['f2'];
      log('采购订单已载入内存，已重置后续步骤缓存', 'success');
    }
    if (idKey === 'f-main-s2') {
      delete resultsMemoryRef.current['f2'];
      log('手动上传了门店文件，Step 2 将优先使用此文件', 'warn');
    }
    if (idKey === 'f-main-s3') {
      delete resultsMemoryRef.current['f3'];
    }
    if (shouldPersist) {
      await saveFileDB(idKey, file);
      setCachedFiles(prev => ({ ...prev, [idKey]: true }));
      log(`已存入记忆库: ${file.name}`, 'success');
      toast.success('文件已缓存', { description: '下次刷新无需重新上传' });
    }
  };

  const renderPreview = (data: PreviewData | null) => setPreviewData(data);

  const executeDownload = () => {
    if (currentDownloadFnRef.current) {
      currentDownloadFnRef.current();
      toast.success('文件导出成功');
    } else {
      showError('没有可导出的数据', '请先执行操作生成结果');
    }
  };

  /* ── Fulu Execute ── */
  const executeFulu = async () => {
    setIsProcessing(true);
    setShowLog(true);
    try {
      if (currentStep === 1) await runFulu1();
      else if (currentStep === 2) await runFulu2();
      else await runFulu3();
    } finally { setIsProcessing(false); }
  };

  const runFulu1 = async () => {
    const wbOrder = memPurchaseOrderWBRef.current || await getWorkbook('f-main-s1');
    const wbHist = await getWorkbook('f-hist');
    if (!wbOrder) { showError('请上传主文件：厂家直发采购订单.xlsx'); return; }

    log('== 开始提取门店信息 ==', 'info');
    const dfOrder = XLSX.utils.sheet_to_json(wbOrder.Sheets[wbOrder.SheetNames[0]], { defval: '' });
    const shippedCodes = new Set<string>();
    const samples: string[] = [];

    if (wbHist) {
      log('正在执行深度扫描（全表所有工作表数字串提取）...', 'info');
      let totalRows = 0;
      wbHist.SheetNames.forEach(sheetName => {
        const sheet = wbHist.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
        data.forEach(row => {
          totalRows++;
          row.forEach(cell => {
            if (!cell) return;
            const text = String(cell);
            const matches = text.match(/\d{6,12}/g);
            if (matches) {
              matches.forEach(m => {
                const c6 = m.slice(-6);
                shippedCodes.add(c6);
                if (samples.length < 10 && !samples.includes(c6)) samples.push(c6);
              });
            }
          });
        });
      });
      log(`扫描完成！共处理 ${totalRows} 行，识别出 ${shippedCodes.size} 个历史编号`, 'success');
      if (samples.length > 0) log(`历史编号示例 (后6位): ${samples.join(', ')}`, 'info');
    }

    const EXPECTED: Record<string, string> = {
      'F360010': '收银机', 'F360011': '小票打印机（蓝牙）', 'F360013': '扫码盒子',
      'F360014': '钱箱', 'F360015': '电子菜单屏', 'F360022': '杯贴打印机（235B）'
    };

    const groups: Record<string, any[]> = {};
    (dfOrder as any[]).forEach((r: any) => {
      const n = String(r['门店名称'] || '').trim();
      if (n && n !== 'undefined') {
        if (!groups[n]) groups[n] = [];
        groups[n].push(r);
      }
    });

    const results: StoreInfo[] = [];
    let errCount = 0;

    Object.entries(groups).forEach(([storeName, group]) => {
      const products: Record<string, string> = {};
      group.forEach(row => {
        const code = String(row['物料编码'] || '').trim();
        const name = String(row['物料名称'] || '').trim();
        if (code) products[code] = name;
      });

      const missing: string[] = [];
      Object.entries(EXPECTED).forEach(([code, name]) => {
        if (!products[code]) missing.push(`${name}(${code})`);
      });

      if (Object.keys(products).length === 6 && missing.length === 0) {
        const info = group[0];
        const storeCode = String(info['门店编码'] || '').trim();
        const rawAddr = String(info['门店收货地址'] || '').trim();
        const fullAddr = completeAddress(
          rawAddr,
          info['收货省'] || info['省'] || '',
          info['收货市'] || info['市'] || '',
          info['收货区'] || info['区'] || ''
        );
        const normalizedStoreCode = storeCode.slice(-6);
        const isDup = shippedCodes.has(normalizedStoreCode);
        results.push({ storeName, storeCode, address: fullAddr.replace(/^市辖区/, ''), contact: String(info['门店收货人'] || '').trim(), phone: String(info['门店收货人电话'] || '').trim(), _isDup: isDup });
      } else { errCount++; }
    });

    const stats = `符合条件: ${results.length} 家 | 其中重复: ${results.filter(r => r._isDup).length} 家 | 跳过: ${errCount} 家`;
    const viewRows = results.map(r => ({ '状态': r._isDup ? '🔴 重复' : '🟢 新店', '门店名称': r.storeName, '门店编号': r.storeCode, '地址': r.address, '姓名': r.contact, '电话': r.phone, _isDup: r._isDup }));
    const pd: PreviewData = { title: '提取门店预览', headers: ['状态', '门店名称', '门店编号', '姓名', '电话', '地址'], rows: viewRows, stats };

    const downloadFn = () => {
      const ws = XLSX.utils.aoa_to_sheet(results.map(r => [`门店名称：${r.storeName}\n门店编号：${r.storeCode}\n门店地址：${r.address}\n姓名：${r.contact}\n联系方式：${r.phone}`]));
      ws['!cols'] = [{ wch: 60 }];
      for (let i = 0; i < results.length; i++) {
        const cell = ws[XLSX.utils.encode_cell({ r: i, c: 0 })];
        cell.s = { alignment: { wrapText: true, vertical: 'center' } };
        if (results[i]._isDup) cell.s.font = { color: { rgb: 'FF0000' }, bold: true };
      }
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      const filename = `提取福鹿家门店${getDateStr()}共${results.length}单.xlsx`;
      XLSX.writeFile(wb, filename);
    };

    memExtractedStoresRef.current = results;
    resultsMemoryRef.current['f1'] = { preview: pd, download: downloadFn, filename: `提取福鹿家门店${getDateStr()}共${results.length}单.xlsx` };
    renderPreview(pd);
    currentDownloadFnRef.current = downloadFn;

    const dupCount1 = results.filter(r => r._isDup).length;
    if (dupCount1 > 0) showDupWarning(dupCount1, `第1步：${dupCount1} 家门店与历史发货记录重复`);
    else toast.success(`提取完成，共 ${results.length} 家门店，无重复`);
    log('提取完成', 'success');
  };

  const runFulu2 = async () => {
    log('== 开始准备生成订单数据 ==', 'info');
    let stores: StoreInfo[] | null = null;

    // 1. 优先检查：是否手动上传了文件？
    const manualFile = uploadedFiles['f-main-s2'];
    if (manualFile) {
      log(`[手动文件] 检测到已上传文件: ${manualFile.name}，正在解析...`, 'info');
      const wb = await getWorkbook('f-main-s2');
      if (wb) {
        stores = [];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) as any[][];
        raw.forEach(r => {
          const text = String(r[0] || '');
          if (!text.includes('门店名称')) return;
          const info: Record<string, string> = {};
          text.split('\n').forEach(l => {
            const p = l.split(/[:：]/);
            if (p.length >= 2) info[p[0].trim()] = p.slice(1).join(':').trim();
          });
          if (info['联系方式'] || info['电话']) {
            stores!.push({
              storeName: info['门店名称'], storeCode: info['门店编号'],
              address: info['门店地址'] || info['地址'], contact: info['姓名'],
              phone: info['联系方式'] || info['电话'],
              _isDup: text.includes('重复') || text.includes('已发过货')
            });
          }
        });
        if (stores.length > 0) {
          log(`[解析成功] 从文件中识别出 ${stores.length} 家门店`, 'success');
        } else {
          log('[解析失败] 上传的文件格式不正确', 'error');
          stores = null;
        }
      }
    }

    // 2. 如果没上传文件，再尝试从内存获取第一步的实时结果
    if (!stores) {
      stores = memExtractedStoresRef.current;
      if (stores && stores.length > 0) {
        log(`[内存复用] 检测到第 1 步生成的 ${stores.length} 家门店结果（首店：${stores[0].storeName}），直接使用`, 'success');
      }
    }

    if (!stores || stores.length === 0) {
      showError('无法获取门店数据', '请先执行第 1 步，或者在第 2 步手动上传文件');
      return;
    }
    const wbTpl = await getWorkbook('f-tpl');
    if (!wbTpl) { showError('缺少 订单分表模版.xlsx'); return; }

    log('== 开始生成吉客云订单 ==', 'info');
    const dfProd = XLSX.utils.sheet_to_json(wbTpl.Sheets['货品']);
    const pMap: Record<string, any> = {};
    (dfProd as any[]).forEach((p: any) => (pMap[p['货品名称']] = p));

    const FIXED_PRODS = ['（上海商米）D3PRO单屏', 'XP-80T（USB+蓝牙）', 'XP-235B（USB）', 'XL-2330支付盒子', 'JY-335C黑钱箱', '惠科电子菜单屏（代发）'];
    const DEF: Record<string, string> = { '物流公司': '德邦特惠', '业务员': '王德龙', '客户账号': '鲜啤福鹿家', '销售渠道名称': '仝心科技线下批发', '结算方式': '欠款计应收' };

    const orderRows: any[] = [];
    const productRows: any[] = [];
    const phoneCounter: Record<string, number> = {};
    let totalAmt = 0;

    stores.forEach(s => {
      phoneCounter[s.phone] = (phoneCounter[s.phone] || 0) + 1;
      let amt = 0;
      const pNames: string[] = [];
      FIXED_PRODS.forEach(name => {
        const p = pMap[name];
        if (p) {
          amt += p['单价'];
          productRows.push({ '导入编号(关联订单)': s.phone, '货品名称': p['货品名称'], '条码': p['条码'] || '', '货品编号': p['货品编号'] || '', '规格': p['规格'] || '', '数量': 1, '单价': p['单价'] });
          pNames.push(`${name}*1`);
        }
      });
      totalAmt += amt;
      const isDup = s._isDup || phoneCounter[s.phone] > 1;
      orderRows.push({ '导入编号': s.phone, '收货人': '', '手机': '', '收货地址': '', '收货人信息(解析)': `${s.contact}，${s.phone}，${s.address}`, '应收邮资': 0, '应收合计': amt, '客服备注': `${pNames.join('+')} （${s.storeName} 门店编号：${s.storeCode}）`, ...DEF, _isDup: isDup });
    });
    orderRows.forEach(r => { if (phoneCounter[r['导入编号']] > 1) r._isDup = true; });

    // Step 2 also flags rows where same phone appears more than once
    // (phoneCounter already set above; refresh _isDup for any that slipped through)
    orderRows.forEach(r => {
      if ((phoneCounter[r['导入编号']] || 0) > 1) r._isDup = true;
    });

    const dupCount2 = orderRows.filter(r => r._isDup).length;
    const pd: PreviewData = { title: '生成订单预览', headers: ['导入编号', '收货人信息(解析)', '应收合计', '客服备注', '物流公司'], rows: orderRows, stats: `生成订单: ${orderRows.length} 条 | 总金额: ￥${totalAmt.toFixed(2)} | 重复: ${dupCount2} 条` };

    const downloadFn = () => {
      const wb = XLSX.utils.book_new();
      const wsOrder = XLSX.utils.json_to_sheet(orderRows.map(r => { const o: any = {}; Object.keys(r).forEach(k => { if (!k.startsWith('_')) o[k] = r[k]; }); return o; }));
      wsOrder['!cols'] = Array(15).fill({ wch: 20 });
      orderRows.forEach((r, i) => {
        if (r._isDup) {
          const range = XLSX.utils.decode_range(wsOrder['!ref']!);
          for (let c = range.s.c; c <= range.e.c; c++) { const cell = wsOrder[XLSX.utils.encode_cell({ r: i + 1, c })]; if (cell) cell.s = { font: { color: { rgb: 'FF0000' }, bold: true } }; }
        }
      });
      XLSX.utils.book_append_sheet(wb, wsOrder, '订单');
      const wsProd = XLSX.utils.json_to_sheet(productRows);
      wsProd['!cols'] = Array(10).fill({ wch: 15 });
      XLSX.utils.book_append_sheet(wb, wsProd, '货品');
      XLSX.writeFile(wb, `福鹿家ERP订单${getDateStr()}共${orderRows.length}单.xlsx`);
    };

    resultsMemoryRef.current['f2'] = { preview: pd, download: downloadFn };
    renderPreview(pd);
    currentDownloadFnRef.current = downloadFn;

    if (dupCount2 > 0) showDupWarning(dupCount2, `第2步：${dupCount2} 条订单存在重复电话/门店`);
    else toast.success(`订单生成完成，共 ${orderRows.length} 条，无重复`);
    log('生成完成', 'success');
  };

  const runFulu3 = async () => {
    const wbTarget = await getWorkbook('f-main-s3');
    const wbOrder = memPurchaseOrderWBRef.current || await getWorkbook('f-order');
    const wbJike = await getWorkbook('f-jike');
    const wbMenu = await getWorkbook('f-menu');
    if (!wbTarget || !wbOrder) { showError('缺少主文件或采购订单源'); return; }

    log('== 开始匹配快递单号 ==', 'info');
    const orderMap: Record<string, any> = {};
    XLSX.utils.sheet_to_json(wbOrder.Sheets[wbOrder.SheetNames[0]]).forEach((r: any) => {
      const oms = String(r['OMS厂家直发单号'] || '').trim();
      if (oms) orderMap[oms] = { name: String(r['门店收货人'] || '').trim(), mat: String(r['物料名称'] || '').trim() };
    });

    const jikeMap: Record<string, string> = {};
    if (wbJike) XLSX.utils.sheet_to_json(wbJike.Sheets[wbJike.SheetNames[0]]).forEach((r: any) => { const n = String(r['收货人'] || '').trim(); const t = String(r['物流单号'] || '').trim(); if (n && t && !jikeMap[n]) jikeMap[n] = t; });

    const menuMap: Record<string, string> = {};
    if (wbMenu) {
      const reg1 = /^(.+?)(KY|DPK|SF|YT|ZTO|STO|HTKY|JT[A-Za-z0-9]+)$/i;
      const reg2 = /^(KY|DPK|SF|YT|ZTO|STO|HTKY|JT[A-Za-z0-9]+)(.+)$/i;
      (XLSX.utils.sheet_to_json(wbMenu.Sheets[wbMenu.SheetNames[0]], { header: 1 }) as any[][]).forEach(r => {
        const val = String(r[0] || '').trim();
        let m = val.match(reg1);
        if (m) { menuMap[m[1].trim()] = m[2].trim(); return; }
        m = val.match(reg2);
        if (m) { menuMap[m[2].trim()] = m[1].trim(); return; }
      });
    }

    const parseCo = (no: string) => {
      const n = no.toUpperCase();
      if (n.startsWith('DPK')) return '德邦快递'; if (n.startsWith('KY')) return '跨越速运';
      if (n.startsWith('SF')) return '顺丰速运'; if (n.startsWith('YT')) return '圆通快递';
      if (n.startsWith('ZTO')) return '中通快递'; if (n.startsWith('STO')) return '申通快递';
      if (n.startsWith('JT')) return '极兔快递'; return '其他';
    };

    const wsT = wbTarget.Sheets[wbTarget.SheetNames[0]];
    const range = XLSX.utils.decode_range(wsT['!ref']!);
    const head: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) head[c] = String(wsT[XLSX.utils.encode_cell({ r: range.s.r, c })]?.v || '').trim();

    const idxOms = head.indexOf('OMS厂家直发单号');
    const idxTrack = head.indexOf('快递单号');
    const idxCo = head.indexOf('快递公司');
    if (idxOms === -1 || idxTrack === -1 || idxCo === -1) { showError('目标表缺少必要的列 (OMS/快递单号/快递公司)'); return; }

    // Pre-scan for duplicate OMS numbers in the target file
    const omsSeenInFile: Record<string, number> = {};
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const oms = String(wsT[XLSX.utils.encode_cell({ r, c: idxOms })]?.v || '').trim();
      if (oms) omsSeenInFile[oms] = (omsSeenInFile[oms] || 0) + 1;
    }
    const dupOmsSet = new Set(Object.keys(omsSeenInFile).filter(k => omsSeenInFile[k] > 1));
    if (dupOmsSet.size > 0) {
      dupOmsSet.forEach(oms => log(`[重复OMS] ${oms} 在目标表中出现 ${omsSeenInFile[oms]} 次`, 'error'));
    }

    const viewRows: any[] = [];
    let okCnt = 0; let failCnt = 0;

    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const oms = String(wsT[XLSX.utils.encode_cell({ r, c: idxOms })]?.v || '').trim();
      if (!oms) continue;
      const info = orderMap[oms];
      let track = ''; let co = ''; let status = '未找到';
      if (info) {
        const isMenu = info.mat.includes('电子菜单屏') || String(wsT[XLSX.utils.encode_cell({ r, c: head.indexOf('物料编码') })]?.v).includes('F360015');
        track = isMenu ? menuMap[info.name] : jikeMap[info.name];
        if (track) { co = parseCo(track); status = '匹配成功'; okCnt++; wsT[XLSX.utils.encode_cell({ r, c: idxTrack })] = { v: track, t: 's' }; wsT[XLSX.utils.encode_cell({ r, c: idxCo })] = { v: co, t: 's' }; }
        else { status = isMenu ? '菜单屏缺失' : '吉客云缺失'; failCnt++; log(`[缺失] OMS:${oms} 收货人:${info.name}`, 'warn'); }
      } else { failCnt++; }
      const isDupOms = dupOmsSet.has(oms);
      viewRows.push({ 'OMS单号': oms, '收货人': info ? info.name : '未知', '匹配物流': track, '快递公司': co, '状态': isDupOms ? '⚠️ OMS重复' : status, _isDup: isDupOms, _isWarn: !track && !isDupOms });
    }

    const dupCount3 = dupOmsSet.size;
    const pd: PreviewData = { title: '回填快递预览', headers: ['OMS单号', '收货人', '匹配物流', '快递公司', '状态'], rows: viewRows, stats: `匹配成功: ${okCnt} | 匹配失败: ${failCnt} | OMS重复: ${dupCount3} 个` };

    const downloadFn = () => { XLSX.writeFile(wbTarget, '厂家直发采购快递单_已填写.xlsx'); };

    resultsMemoryRef.current['f3'] = { preview: pd, download: downloadFn };
    renderPreview(pd);
    currentDownloadFnRef.current = downloadFn;

    if (dupCount3 > 0) showDupWarning(dupCount3, `第3步：${dupCount3} 个 OMS单号 在目标表中重复出现`);
    if (failCnt > 0) toast.warning(`${failCnt} 条快递单号未能匹配`, { description: '黄色行需手动补全', duration: 5000 });
    if (dupCount3 === 0 && failCnt === 0) toast.success(`回填完成，${okCnt} 条全部匹配成功`);
    log('匹配完成', 'success');
  };

  /* ── Shangsong ── */
  const processSS = async () => {
    setIsProcessing(true);
    setShowLog(true);
    try {
      const wbMain = await getWorkbook('s-main-input');
      const wbTpl = await getWorkbook('s-tpl');
      const wbPrice = await getWorkbook('s-price');
      if (!wbMain) { showError('请上传待转换的商颂订单'); return; }
      if (!wbTpl) { showError('缺少订单分表模版.xlsx'); return; }

      log('==== 开始执行商颂转换 ====', 'info');
      const dfOrder = XLSX.utils.sheet_to_json((wbTpl.Sheets['订单'] || wbTpl.Sheets[wbTpl.SheetNames[0]]));
      const dfProduct = XLSX.utils.sheet_to_json((wbTpl.Sheets['货品'] || wbTpl.Sheets[wbTpl.SheetNames[1]]));
      const productMap: Record<string, any> = {};
      const printerPaperWidthMap: Record<string, number> = {};
      const paperByWidthMap: Record<number, any> = {};

      (dfProduct as any[]).forEach((row: any) => {
        if (row['货品名称'] && row['单价'] !== undefined) {
          const info = { name: row['货品名称'], code: row['货品编号'] || '', barcode: row['条码'] || '', spec: row['规格'] || '', price: parseFloat(row['单价']), paper_width: parseFloat(row['纸张宽度']) || null };
          productMap[row['货品名称']] = info;
          if (info.paper_width && !row['货品名称'].includes('纸')) printerPaperWidthMap[row['货品名称']] = info.paper_width;
          else if (row['货品名称'].includes('纸') && info.paper_width && !paperByWidthMap[info.paper_width]) paperByWidthMap[info.paper_width] = info;
        }
      });

      const defaults: Record<string, string> = { '业务员': '王德龙', '客户账号': '鲜啤福鹿家', '销售渠道名称': '仝心科技线下批发', '结算方式': '欠款计应收', '物流公司': '中通快递' };
      if ((dfOrder as any[]).length > 0) ['业务员', '物流公司', '客户账号', '销售渠道名称', '结算方式'].forEach(col => { if ((dfOrder as any[])[0][col]) defaults[col] = (dfOrder as any[])[0][col]; });

      const expressRules: Record<string, any> = {};
      const postageRules: Record<string, number> = {};
      if (wbPrice) {
        let curPlat = '淘宝';
        const priceSheet = wbPrice.Sheets[wbPrice.SheetNames[0]];
        const priceRows = XLSX.utils.sheet_to_json(priceSheet, { header: 1 }) as any[][];
        log(`开始同步快递规则，共 ${priceRows.length} 行`, 'info');

        priceRows.forEach((row) => {
          const firstCol = String(row[0] || '').trim();
          if (firstCol.includes('淘宝')) { curPlat = '淘宝'; return; }
          if (firstCol.includes('拼多多') || firstCol.includes('多多')) { curPlat = '拼多多'; return; }
          if (firstCol === '省份' || !firstCol || firstCol === 'nan') return;

          const province = extractProvince(firstCol);
          if (province) {
            const key = `${province}_${curPlat}`;
            const postageRaw = String(row[3] || '').replace('元', '').trim();
            const postageVal = postageRaw === 'nan' ? 0 : (parseFloat(postageRaw) || 0);

            expressRules[key] = {
              '1_2kg': String(row[1] || '').trim(),
              '2kg_plus': String(row[2] || '').trim()
            };
            postageRules[key] = postageVal;
          }
        });
        log(`规则库对齐完毕: ${Object.keys(expressRules).length} 条规则已加载`, 'success');
      }

      const orders: any[] = [];
      const mainSheet = wbMain.Sheets[wbMain.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(mainSheet, { header: 1 }) as any[][];
      rawData.forEach(row => {
        const c0 = String(row[0] || '').trim();
        const c1 = String(row[1] || '').trim();
        const c2 = String(row[2] || '').trim();
        const c3 = String(row[3] || '').trim();
        if (c0.includes('订单') || !c0 || c0 === 'nan') return;

        // 1:1 还原 Python 平台判定逻辑
        const isPDD = /^\d{6}-\d+/.test(c0);
        orders.push({
          id: c0,
          interface: c1,
          address: c2,
          remark: c3 === 'nan' ? '' : c3,
          platform: isPDD ? '拼多多' : '淘宝'
        });
      });

      const normalize = (s: string) => s.replace(/（/g, '(').replace(/）/g, ')').toUpperCase().replace(/\s/g, '');
      const findMatch = (name: string) => {
        if (productMap[name]) return productMap[name];
        const nN = normalize(name);
        for (let k in productMap) { if (normalize(k) === nN || k.includes(name) || name.includes(k)) return productMap[k]; }
        return null;
      };

      const orderRows: any[] = [];
      const productRows: any[] = [];
      let sumTotal = 0; let sumPost = 0;

      orders.forEach(order => {
        const lines = order.interface.split('\n');
        const mList: any[] = [];
        const ips: string[] = [];
        let pTotal = 0;
        const prnByW: Record<number, number> = {};
        const pprByW: Record<number, number> = {};

        lines.forEach((line: string) => {
          line = line.trim();
          if (!line || ['送网线', '网线', '收据', '测试纸', '顺丰到付'].some(f => line.includes(f))) return;
          const ipM = line.match(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/g) || [];
          ips.push(...ipM);
          const clean = line.replace(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/g, '').trim();
          const m = clean.match(/^(.+?)\s*\*\s*(\d+)\s*$/);
          const name = m ? m[1].trim() : clean;
          const qty = m ? parseInt(m[2]) : 1;
          const prod = findMatch(name);
          if (prod) {
            const price = prod.price * qty;
            pTotal += price; sumTotal += price;
            mList.push({ name: prod.name, price: prod.price, qty });
            productRows.push({ '导入编号(关联订单)': order.id, '货品名称': prod.name, '条码': prod.barcode, '货品编号': prod.code, '规格': prod.spec, '数量': qty, '单价': prod.price });
            // 1:1 还原 Python 补纸统计逻辑
            if (prod.price > 50 && !prod.name.includes('纸')) {
              const w = printerPaperWidthMap[prod.name] || 80;
              prnByW[w] = (prnByW[w] || 0) + qty;
            }
            else if (prod.name.includes('纸') && prod.paper_width) {
              pprByW[prod.paper_width] = (pprByW[prod.paper_width] || 0) + qty;
            }
          } else { log(`未找到: ${name}`, 'error'); }
        });

        // 补纸处理
        for (let w in prnByW) {
          const toAdd = Math.max(0, prnByW[w] - (pprByW[w] || 0));
          const pPaper = paperByWidthMap[w];
          if (toAdd > 0 && pPaper) {
            log(`[补纸] 订单 ${order.id.slice(-6)} 补 ${pPaper.name}*${toAdd}`, 'warn');
            const price = pPaper.price * toAdd; pTotal += price; sumTotal += price;
            mList.push({ name: pPaper.name, qty: toAdd });
            productRows.push({ '导入编号(关联订单)': order.id, '货品名称': pPaper.name, '条码': pPaper.barcode, '货品编号': pPaper.code, '规格': pPaper.spec, '数量': toAdd, '单价': pPaper.price });
          }
        }

        const prov = extractProvince(order.address);
        let expCo = defaults['物流公司'] || '德邦特惠';
        let post = 0;

        if (prov && wbPrice) {
          // 1:1 还原 Python 重量判定：所有单价 > 50 的商品总数
          const printerCount = mList.reduce((acc, p) => acc + (p.price > 50 ? p.qty : 0), 0);
          const weightClass = printerCount >= 2 ? '2kg_plus' : '1_2kg';
          const ruleKey = `${prov}_${order.platform}`;

          if (expressRules[ruleKey]) {
            expCo = expressRules[ruleKey][weightClass];
            post = postageRules[ruleKey] || 0;
            log(`  [${order.platform}] ${order.id.slice(-8)} | ${prov} | ${weightClass === '2kg_plus' ? '重货' : '普货'} | ${expCo} | 邮资:+￥${post}`, 'success');
          } else {
            log(`  [匹配失败] ${order.id.slice(-8)} | 省份:${prov} 平台:${order.platform} 未找到快递规则`, 'warn');
          }
        } else if (!prov) {
          log(`  [识别失败] ${order.id.slice(-8)} | 地址无法识别省份: ${order.address.slice(0, 10)}...`, 'error');
        }

        sumPost += post;
        const mgMap: Record<string, number> = {};
        mList.forEach(p => (mgMap[p.name] = (mgMap[p.name] || 0) + p.qty));
        let rem = Object.entries(mgMap).map(([n, q]) => `${n}*${q}`).join('+');
        if (ips.length) rem += ` 改IP：${[...new Set(ips)].join(',')}`;
        if (order.remark) rem += ` ${order.remark}`;

        orderRows.push({
          ...defaults,
          '导入编号': order.id, '收货人': '', '手机': '', '收货地址': '',
          '收货人信息(解析)': order.address, '应收邮资': post,
          '应收合计': pTotal + post, '客服备注': rem, '物流公司': expCo,
          _isWarn: !prov || !expCo
        });
      });

      // Check for duplicate order IDs in 商颂
      const ssIdCounter: Record<string, number> = {};
      orders.forEach(o => { ssIdCounter[o.id] = (ssIdCounter[o.id] || 0) + 1; });
      orderRows.forEach(r => {
        if ((ssIdCounter[r['导入编号']] || 0) > 1) {
          r._isDup = true;
          log(`[重复订单] ${r['导入编号']} 出现 ${ssIdCounter[r['导入编号']]} 次`, 'error');
        }
      });

      const dupCountSS = orderRows.filter(r => r._isDup).length;
      const pd: PreviewData = { title: '商颂订单预览', headers: ['导入编号', '收货人信息(解析)', '应收邮资', '应收合计', '客服备注', '物流公司'], rows: orderRows, stats: `总单: ${orders.length} | 货品: ￥${sumTotal.toFixed(2)} | 邮资: ￥${sumPost.toFixed(2)} | 重复: ${dupCountSS} 条` };

      const downloadFn = () => {
        const wb = XLSX.utils.book_new();
        
        // 动态提取模板中的原始列顺序
        const tplSheet = wbTpl.Sheets[wbTpl.SheetNames[0]];
        const tplHeaders = (XLSX.utils.sheet_to_json(tplSheet, { header: 1 })[0] as string[]) || [];
        
        // 确保我们的核心字段都在里面
        const exportHeaders = tplHeaders.length > 0 ? tplHeaders : ['导入编号', '收货人', '手机', '收货地址', '应收邮资', '应收合计', '客服备注', '物流公司', '收货人信息(解析)', '业务员', '客户账号', '销售渠道名称', '结算方式'];

        const wsOrder = XLSX.utils.json_to_sheet(orderRows.map(r => {
          const o: any = {};
          Object.keys(r).forEach(k => { if (!k.startsWith('_')) o[k] = r[k]; });
          return o;
        }), { header: exportHeaders });

        wsOrder['!cols'] = Array(Math.max(15, exportHeaders.length)).fill({ wch: 18 });
        XLSX.utils.book_append_sheet(wb, wsOrder, '订单');
        const wsProd = XLSX.utils.json_to_sheet(productRows);
        wsProd['!cols'] = Array(8).fill({ wch: 15 });
        XLSX.utils.book_append_sheet(wb, wsProd, '货品');
        XLSX.writeFile(wb, `商颂${getDateStr()}共${orders.length}单.xlsx`);
      };

      resultsMemoryRef.current['ss'] = { preview: pd, download: downloadFn };
      renderPreview(pd);
      currentDownloadFnRef.current = downloadFn;

      if (dupCountSS > 0) showDupWarning(dupCountSS, `商颂：${dupCountSS} 条订单编号重复出现`);
      else toast.success(`转换完成，共 ${orders.length} 单，无重复`);
      log('转换完成', 'success');
    } finally { setIsProcessing(false); }
  };

  /* ── Step Configs ── */
  const stepConfigs: Record<number, { title: string; desc: string; main: string; subs: { id: string; label: string }[] }> = {
    1: { title: '提取门店信息', desc: '过滤重复项并提取合格门店', main: '厂家直发采购订单.xlsx', subs: [{ id: 'f-hist', label: '福鹿家总发货记录.xlsx（查重）' }] },
    2: { title: '生成 ERP 订单', desc: '根据提取的门店生成分表', main: '福鹿家_提取后门店.xlsx', subs: [{ id: 'f-tpl', label: '订单分表模板.xlsx（缓存）' }] },
    3: { title: '回填快递单号', desc: '匹配并回填快递物流信息', main: '厂家直发采购快递单.xlsx', subs: [{ id: 'f-order', label: '厂家直发采购订单.xlsx' }, { id: 'f-jike', label: '快递单号-吉客云.xlsx' }, { id: 'f-menu', label: '快递单号-菜单屏.xlsx' }] },
  };

  const handleStepClick = (step: number) => {
    setCurrentStep(step);
    setExpandedStep(expandedStep === step ? null : step);

    // 关键修复：从记忆库中提取 .preview，避免白屏
    const key = `f${step}`;
    const mem = resultsMemoryRef.current[key];
    if (mem) {
      renderPreview(mem.preview);
      currentDownloadFnRef.current = mem.download;
    } else {
      renderPreview(null);
      currentDownloadFnRef.current = null;
    }
  };

  useEffect(() => {
    let memoryKey = '';
    if (activeTab === 'fulujia') {
      memoryKey = `f${currentStep}`;
    } else if (activeTab === 'shangsong') {
      memoryKey = 'ss';
    }

    const mem = resultsMemoryRef.current[memoryKey];
    if (mem) {
      renderPreview(mem.preview);
      currentDownloadFnRef.current = mem.download;
    } else {
      renderPreview(null);
      currentDownloadFnRef.current = null;
    }
  }, [activeTab, currentStep]);

  /* ── File Upload Zone ── */
  const FileZone = ({ idKey, label, mainStyle = false, persistOnUpload = false, autoName }: { idKey: string; label: string; mainStyle?: boolean; persistOnUpload?: boolean; autoName?: string }) => {
    const hasFile = !!(uploadedFiles[idKey] || cachedFiles[idKey] || autoName);
    return (
      <label className={`
        group relative flex items-center gap-3 rounded-xl border cursor-pointer transition-all duration-200 select-none h-full p-3
        ${hasFile
          ? (isDark ? 'border-green-500/40 bg-green-500/10' : 'border-green-500/30 bg-green-50')
          : (isDark ? 'border-white/8 bg-white/3 hover:bg-white/6' : 'border-black/8 bg-gray-50/50 hover:bg-gray-100/70')
        }
      `}>
        <input type="file" accept=".xlsx,.xls" className="hidden"
          onChange={e => { if (e.target.files?.[0]) handleFileUpload(e.target.files[0], idKey, persistOnUpload); }} />
        <div className={`shrink-0 w-8 h-8 rounded-lg flex items-center justify-center ${hasFile ? 'bg-green-500/15' : (isDark ? 'bg-white/6' : 'bg-gray-200')}`}>
          {hasFile ? <FileCheck className="w-4 h-4 text-green-400" /> : <Upload className={`w-4 h-4 ${isDark ? 'text-gray-500' : 'text-gray-400'}`} />}
        </div>
        <div className="flex-1 min-w-0">
          <p className={`text-xs font-medium truncate ${isDark ? 'text-gray-300' : 'text-gray-600'}`}>{label}</p>
          <p className={`text-xs truncate mt-0.5 ${hasFile ? 'text-green-400' : (isDark ? 'text-gray-500' : 'text-gray-400')}`}>
            {uploadedFiles[idKey]?.name || (cachedFiles[idKey] ? '已缓存，可复用' : (autoName ? `自动加载: ${autoName}` : '未上传'))}
          </p>
        </div>
        {(persistOnUpload || autoName) && (
          <span className={`text-[10px] px-1.5 py-0.5 rounded-full ${autoName ? (isDark ? 'bg-green-500/15 text-green-400' : 'bg-green-100 text-green-600') : (isDark ? 'bg-blue-500/15 text-blue-400' : 'bg-blue-100 text-blue-600')}`}>
            {autoName ? '自动' : '缓存'}
          </span>
        )}
      </label>
    );
  };

  /* ═════════════════════ RENDER ═════════════════════ */
  return (
    <>
      <Toaster richColors position="top-center" />

      {/* ── Error Modal (center screen) ── */}
      {errorModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{ background: 'rgba(0,0,0,0.55)', backdropFilter: 'blur(12px)' }}
          onClick={() => setErrorModal(null)}
        >
          <div
            className={`relative w-full max-w-sm rounded-2xl shadow-2xl border p-6 animate-in fade-in zoom-in-95 duration-200
              ${isDark ? 'bg-[#2c2c2e] border-white/12' : 'bg-white border-black/8'}`}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex items-start gap-4">
              <div className="shrink-0 w-11 h-11 rounded-full bg-red-500/15 flex items-center justify-center">
                <AlertTriangle className="w-6 h-6 text-red-500" />
              </div>
              <div className="flex-1 min-w-0 pt-0.5">
                <h3 className={`font-semibold leading-snug ${isDark ? 'text-white' : 'text-gray-900'}`}>{errorModal.title}</h3>
                {errorModal.desc && <p className={`text-sm mt-1.5 leading-relaxed ${isDark ? 'text-gray-400' : 'text-gray-500'}`}>{errorModal.desc}</p>}
              </div>
              <button onClick={() => setErrorModal(null)} className={`shrink-0 w-7 h-7 rounded-full flex items-center justify-center transition-colors ${isDark ? 'hover:bg-white/10 text-gray-500' : 'hover:bg-gray-100 text-gray-400'}`}>
                <X className="w-4 h-4" />
              </button>
            </div>
            <button
              onClick={() => setErrorModal(null)}
              className="mt-5 w-full py-2.5 rounded-xl bg-[#007AFF] hover:bg-[#0071e3] text-white text-sm font-semibold transition-colors"
            >
              好的
            </button>
          </div>
        </div>
      )}

      {/* ── Log Drawer ── */}
      {showLog && (
        <div
          className="fixed inset-0 z-40 flex items-end justify-center p-4"
          style={{ background: 'rgba(0,0,0,0.45)', backdropFilter: 'blur(8px)' }}
          onClick={() => setShowLog(false)}
        >
          <div
            className={`w-full max-w-3xl rounded-2xl border shadow-2xl overflow-hidden animate-in slide-in-from-bottom-4 duration-300
              ${isDark ? 'bg-[#1a1a1a] border-white/10' : 'bg-gray-900 border-white/10'}`}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex items-center justify-between px-5 py-3 border-b border-white/8">
              <div className="flex items-center gap-2">
                <Terminal className="w-4 h-4 text-green-400" />
                <span className="text-xs font-semibold text-gray-300 tracking-wider uppercase">执行终端</span>
                <span className="text-xs text-gray-600">· {logs.length} 条记录</span>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => setLogs([])} className="text-gray-600 hover:text-gray-300 transition-colors p-1">
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
                <button onClick={() => setShowLog(false)} className="text-gray-600 hover:text-gray-300 transition-colors p-1">
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>
            <div ref={logContainerRef} className="h-72 overflow-y-auto p-5 font-mono text-xs space-y-1.5">
              {logs.length === 0 ? (
                <p className="text-gray-600 italic">等待执行操作...</p>
              ) : (
                logs.map((entry, i) => (
                  <div key={i} dangerouslySetInnerHTML={{ __html: logHtml(entry) }} />
                ))
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Main Layout ── */}
      <div className={`min-h-screen ${T.bg(isDark)} transition-colors duration-300`} style={{ fontFamily: '-apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", sans-serif' }}>

        {/* ── Navbar ── */}
        <nav className={`sticky top-0 z-30 border-b ${isDark ? 'bg-[#1c1c1e]/90 border-white/8' : 'bg-white/90 border-black/8'}`}
          style={{ backdropFilter: 'blur(20px)' }}>
          <div className="max-w-screen-2xl mx-auto px-6 h-14 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-7 h-7 rounded-lg bg-[#007AFF] flex items-center justify-center shadow-lg shadow-blue-500/30">
                <span className="text-white text-xs font-bold">FT</span>
              </div>
              <span className={`font-semibold tracking-tight ${isDark ? 'text-white' : 'text-gray-900'}`}>FTH 订单自动化助手</span>
              <span className={`text-xs px-2 py-0.5 rounded-full ${isDark ? 'bg-white/8 text-gray-400' : 'bg-gray-100 text-gray-500'}`}>v3.0</span>
            </div>
            <div className="flex items-center gap-3">
              {/* Log button */}
              <button
                onClick={() => setShowLog(true)}
                className={`relative flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors
                  ${isDark ? 'bg-white/6 hover:bg-white/10 text-gray-300' : 'bg-gray-100 hover:bg-gray-200 text-gray-600'}`}
              >
                <Terminal className="w-3.5 h-3.5" />
                终端
                {logs.length > 0 && <span className="absolute -top-1 -right-1 w-4 h-4 bg-[#007AFF] rounded-full text-[9px] text-white flex items-center justify-center">{logs.length > 99 ? '99' : logs.length}</span>}
              </button>
              {/* Theme toggle */}
              <button
                onClick={() => setIsDark(!isDark)}
                className={`w-9 h-9 rounded-xl flex items-center justify-center transition-all duration-300
                  ${isDark ? 'bg-white/8 hover:bg-white/14 text-yellow-300' : 'bg-gray-100 hover:bg-gray-200 text-gray-600'}`}
              >
                {isDark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
              </button>
            </div>
          </div>
        </nav>

        <div className="max-w-screen-2xl mx-auto px-6 py-6 space-y-5">

          {/* ── Tab Bar ── */}
          <div className={`inline-flex items-center gap-1 p-1 rounded-xl ${isDark ? 'bg-white/6' : 'bg-black/5'}`}>
            {([['fulujia', '📦 福鹿家系列', '#007AFF'], ['shangsong', '🚀 商颂下单', '#ff9500']] as const).map(([id, label, color]) => (
              <button
                key={id}
                onClick={() => setActiveTab(id)}
                className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all duration-200
                  ${activeTab === id
                    ? 'text-white shadow-lg'
                    : (isDark ? 'text-gray-400 hover:text-gray-200' : 'text-gray-500 hover:text-gray-700')
                  }`}
                style={activeTab === id ? { background: color, boxShadow: `0 4px 12px ${color}40` } : {}}
              >
                {label}
              </button>
            ))}
          </div>

          {/* ── Top Collapsible Toolbar ── */}
          {activeTab === 'fulujia' ? (
            <div className={`rounded-2xl border overflow-hidden ${isDark ? 'bg-[#2c2c2e] border-white/8' : 'bg-white border-black/8'} shadow-sm`}>
              {/* Step Pills Row */}
              <div className={`flex items-center gap-2 px-5 py-3 border-b ${isDark ? 'border-white/6 bg-white/2' : 'border-black/5 bg-gray-50/60'}`}>
                <span className={`text-xs font-semibold mr-2 ${isDark ? 'text-gray-500' : 'text-gray-400'}`}>步骤</span>
                {[1, 2, 3].map(step => {
                  const cfg = stepConfigs[step];
                  const isActive = currentStep === step;
                  const isOpen = expandedStep === step;
                  return (
                    <button
                      key={step}
                      onClick={() => handleStepClick(step)}
                      className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-medium transition-all duration-200 border
                        ${isActive
                          ? (isDark ? 'bg-[#007AFF]/15 border-[#007AFF]/40 text-[#007AFF]' : 'bg-blue-50 border-blue-200 text-blue-600')
                          : (isDark ? 'bg-white/4 border-white/6 text-gray-400 hover:bg-white/8 hover:text-gray-200' : 'bg-gray-50 border-black/6 text-gray-500 hover:bg-gray-100 hover:text-gray-700')
                        }`}
                    >
                      <span className={`w-5 h-5 rounded-full flex items-center justify-center text-xs font-bold shrink-0
                        ${isActive ? 'bg-[#007AFF] text-white' : (isDark ? 'bg-white/10 text-gray-500' : 'bg-gray-200 text-gray-500')}`}>
                        {step}
                      </span>
                      <span className="hidden sm:inline">{cfg.title}</span>
                      {isOpen ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                    </button>
                  );
                })}
                <div className="flex-1" />
                {/* Execute */}
                <button
                  onClick={executeFulu}
                  disabled={isProcessing}
                  className="flex items-center gap-2 px-5 py-2 rounded-xl text-sm font-semibold text-white transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed shadow-md"
                  style={{ background: isProcessing ? '#555' : '#007AFF', boxShadow: isProcessing ? 'none' : '0 4px 12px rgba(0,122,255,0.4)' }}
                >
                  {isProcessing ? <><Loader2 className="w-4 h-4 animate-spin" />处理中...</> : <>执行第 {currentStep} 步</>}
                </button>
              </div>

              {/* Expanded Panel */}
              {expandedStep !== null && (() => {
                const subsCount = stepConfigs[expandedStep].subs.length;
                return (
                  <div className="px-8 py-6 animate-in fade-in slide-in-from-top-2 duration-200">
                    {/* Step indicator */}
                    <div className="pb-4 flex items-center justify-center gap-2">
                      <span className={`text-sm font-semibold ${isDark ? 'text-gray-300' : 'text-gray-600'}`}>
                        步骤 {expandedStep}：{stepConfigs[expandedStep].desc}
                      </span>
                      {expandedStep === 2 && memExtractedStoresRef.current && (
                        <span className="flex items-center gap-1 text-xs text-green-400 bg-green-500/10 px-2 py-0.5 rounded-full">
                          <CheckCircle2 className="w-3 h-3" />第1步数据已在内存中
                        </span>
                      )}
                    </div>

                    {/* Left-Right Layout Container */}
                    <div className="max-w-5xl mx-auto">
                      <div className="grid grid-cols-2 gap-4">
                        {/* Left Column - Main File */}
                        <div className="flex-1">
                          <FileZone
                            idKey={`f-main-s${expandedStep}`}
                            label={(expandedStep === 2 && resultsMemoryRef.current['f1']?.filename) ? resultsMemoryRef.current['f1'].filename : stepConfigs[expandedStep].main}
                            autoName={(expandedStep === 2 && !uploadedFiles['f-main-s2']) ? resultsMemoryRef.current['f1']?.filename : undefined}
                            mainStyle
                            persistOnUpload={false}
                          />
                        </div>

                        {/* Right Column - Sub Files */}
                        {subsCount > 0 && (
                          <div className="flex flex-col gap-2">
                            {stepConfigs[expandedStep].subs.map(sub => (
                              <div key={sub.id} className="flex-1">
                                <FileZone idKey={sub.id} label={sub.label} persistOnUpload />
                              </div>
                            ))}
                          </div>
                        )}

                        {/* If no subs, center the main file across both columns */}
                        {subsCount === 0 && <div />}
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          ) : (
            /* Shangsong Toolbar */
            <div className={`rounded-2xl border overflow-hidden ${isDark ? 'bg-[#2c2c2e] border-white/8' : 'bg-white border-black/8'} shadow-sm`}>
              <div className={`flex items-center gap-3 px-5 py-3 border-b ${isDark ? 'border-white/6 bg-white/2' : 'border-black/5 bg-gray-50/60'}`}>
                <span className={`text-xs font-semibold ${isDark ? 'text-gray-400' : 'text-gray-600'}`}>🚀 商颂订单转换</span>
                <button
                  onClick={() => setExpandedSS(!expandedSS)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors
                    ${isDark ? 'bg-white/6 hover:bg-white/10 text-gray-300' : 'bg-gray-100 hover:bg-gray-200 text-gray-600'}`}
                >
                  上传文件 {expandedSS ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                </button>
                <div className="flex-1" />
                <button
                  onClick={processSS}
                  disabled={isProcessing}
                  className="flex items-center gap-2 px-5 py-2 rounded-xl text-sm font-semibold text-white transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed"
                  style={{ background: isProcessing ? '#555' : '#ff9500', boxShadow: isProcessing ? 'none' : '0 4px 12px rgba(255,149,0,0.4)' }}
                >
                  {isProcessing ? <><Loader2 className="w-4 h-4 animate-spin" />处理中...</> : <>开始转换生成</>}
                </button>
              </div>
              {expandedSS && (
                <div className="px-8 py-6 animate-in fade-in slide-in-from-top-2 duration-200">
                  {/* Left-Right Layout Container */}
                  <div className="max-w-5xl mx-auto">
                    <div className="grid grid-cols-2 gap-4">
                      {/* Left Column - Main File */}
                      <div className="flex-1">
                        <FileZone idKey="s-main-input" label="待转换订单 (商颂.xlsx)" mainStyle persistOnUpload={false} />
                      </div>

                      {/* Right Column - Sub Files (2 items stacked) */}
                      <div className="flex flex-col gap-2">
                        <div className="flex-1">
                          <FileZone idKey="s-tpl" label="订单分表模版.xlsx" persistOnUpload />
                        </div>
                        <div className="flex-1">
                          <FileZone idKey="s-price" label="快递价格明细表.xls" persistOnUpload />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ── Preview Area ── */}
          <div className={`rounded-2xl border shadow-sm overflow-hidden
            ${isDark ? 'bg-[#2c2c2e] border-white/8' : 'bg-white border-black/8'}`}>
            {previewData ? (
              <>
                {/* Preview Header */}
                <div className={`flex items-center justify-between px-6 py-4 border-b ${isDark ? 'border-white/6' : 'border-black/6'}`}>
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-xl bg-green-500/15 flex items-center justify-center">
                      <CheckCircle2 className="w-4 h-4 text-green-400" />
                    </div>
                    <div>
                      <h2 className={`font-semibold ${isDark ? 'text-white' : 'text-gray-900'}`}>{previewData.title}</h2>
                      <p className={`text-xs mt-0.5 font-mono ${isDark ? 'text-gray-400' : 'text-gray-500'}`}>{previewData.stats}</p>
                    </div>
                  </div>
                  <button
                    onClick={executeDownload}
                    className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-semibold text-white transition-all"
                    style={{ background: '#34c759', boxShadow: '0 4px 12px rgba(52,199,89,0.35)' }}
                  >
                    <Download className="w-4 h-4" />
                    导出 Excel
                  </button>
                </div>

                {/* Table */}
                <div className="overflow-auto" style={{ maxHeight: 'calc(100vh - 340px)', minHeight: '400px' }}>
                  <table className="w-full text-sm border-collapse">
                    <thead className={`sticky top-0 z-10 ${isDark ? 'bg-[#2c2c2e]' : 'bg-white'}`}>
                      <tr className={`border-b ${isDark ? 'border-white/8' : 'border-black/6'}`}>
                        <th className={`px-5 py-3 text-left text-xs font-semibold w-10 ${isDark ? 'text-gray-500' : 'text-gray-400'}`}>#</th>
                        {previewData.headers.map(h => (
                          <th key={h} className={`px-5 py-3 text-left text-xs font-semibold whitespace-nowrap ${isDark ? 'text-gray-400' : 'text-gray-500'}`}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewData.rows.map((row, i) => (
                        <tr key={i} className={`border-b transition-colors
                          ${row._isDup
                            ? (isDark ? 'bg-red-500/8 border-red-500/10' : 'bg-red-50 border-red-100')
                            : row._isWarn
                              ? (isDark ? 'bg-yellow-500/6 border-yellow-500/10' : 'bg-yellow-50 border-yellow-100')
                              : (isDark ? 'border-white/4 hover:bg-white/3' : 'border-black/4 hover:bg-gray-50')
                          }`}>
                          <td className={`px-5 py-3 text-xs tabular-nums ${isDark ? 'text-gray-600' : 'text-gray-400'}`}>{i + 1}</td>
                          {previewData.headers.map(h => (
                            <td key={h} className={`px-5 py-3 max-w-xs
                              ${row._isDup ? (isDark ? 'text-red-300' : 'text-red-600') : row._isWarn ? (isDark ? 'text-yellow-300' : 'text-yellow-700') : (isDark ? 'text-gray-200' : 'text-gray-700')}`}>
                              <span className="block truncate" title={String(row[h] ?? '')}>{row[h] ?? ''}</span>
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Footer stats bar */}
                <div className={`px-6 py-3 border-t flex items-center gap-4 ${isDark ? 'border-white/6 bg-white/2' : 'border-black/5 bg-gray-50/50'}`}>
                  {previewData.rows.some(r => r._isDup) && (
                    <span className="flex items-center gap-1.5 text-xs text-red-400">
                      <span className="w-2 h-2 rounded-full bg-red-400 inline-block" />
                      红色 = 重复项
                    </span>
                  )}
                  {previewData.rows.some(r => r._isWarn) && (
                    <span className="flex items-center gap-1.5 text-xs text-yellow-400">
                      <span className="w-2 h-2 rounded-full bg-yellow-400 inline-block" />
                      黄色 = 需注意
                    </span>
                  )}
                  <span className={`ml-auto text-xs ${isDark ? 'text-gray-600' : 'text-gray-400'}`}>共 {previewData.rows.length} 行数据</span>
                </div>
              </>
            ) : (
              /* Empty State */
              <div className="flex flex-col items-center justify-center py-28 px-6">
                {/* Modern Icon Stack */}
                <div className="relative mb-6">
                  {/* Background glow */}
                  <div className={`absolute inset-0 rounded-3xl blur-2xl opacity-40 ${isDark ? 'bg-blue-500/20' : 'bg-blue-400/30'}`} />

                  {/* Icon container with gradient */}
                  <div className="relative w-24 h-24 rounded-3xl flex items-center justify-center"
                    style={{
                      background: isDark
                        ? 'linear-gradient(135deg, rgba(0,122,255,0.15) 0%, rgba(52,199,89,0.15) 100%)'
                        : 'linear-gradient(135deg, rgba(0,122,255,0.1) 0%, rgba(52,199,89,0.1) 100%)',
                      border: `1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.06)'}`
                    }}
                  >
                    {/* Stacked icons */}
                    <div className="relative">
                      <FileCheck className={`w-10 h-10 ${isDark ? 'text-blue-400' : 'text-blue-500'}`}
                        style={{ filter: 'drop-shadow(0 2px 8px rgba(0,122,255,0.3))' }} />
                      <Download className={`absolute -bottom-1 -right-1 w-5 h-5 ${isDark ? 'text-green-400' : 'text-green-500'}`}
                        style={{ filter: 'drop-shadow(0 2px 6px rgba(52,199,89,0.4))' }} />
                    </div>
                  </div>
                </div>

                <h3 className={`text-lg font-semibold mb-2 ${isDark ? 'text-gray-200' : 'text-gray-800'}`}>
                  等待数据预览
                </h3>
                <p className={`text-sm text-center max-w-md leading-relaxed ${isDark ? 'text-gray-400' : 'text-gray-500'}`}>
                  上传文件并执行操作后，处理结果将在此处显示
                  <br />
                  <span className={`text-xs ${isDark ? 'text-gray-500' : 'text-gray-400'}`}>
                    支持实时预览、查重标注和一键导出 Excel
                  </span>
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}
