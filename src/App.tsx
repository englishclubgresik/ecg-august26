import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Toaster, toast as sonnerToast } from 'sonner';
import { Users, UserCheck, BookOpen, Calendar as CalendarIcon, DollarSign, FileText, Settings, LogOut, LayoutDashboard, Activity, ChartBar as BarChart3, Plus, Search, ListFilter as Filter, Download, Printer, Share2, Pencil as Edit2, UserCog, Trash2, CircleCheck as CheckCircle2, Circle as XCircle, ChevronDown, Menu, X, SquareCheck as CheckSquare, Briefcase, Bell, CircleAlert as AlertCircle, Eye, RefreshCw, Trash, ArchiveRestore, ArrowLeft, KeyRound, ShieldCheck, Shield, MessageSquare, GraduationCap, Clock, Hash, User, Award, QrCode, Quote, Cloud, CloudOff, Sun, CloudRain, CloudLightning, Droplets, Wind, Thermometer, Link as LinkIcon, MessageCircle, Check, Trophy, Target, Zap, Star, Medal, Mic, Terminal, Copy, Inbox, Database } from 'lucide-react';

// Link Eksekusi Google App Script Anda
const APPSCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwEz5ZxXjcBCQhoyyA0q5VJfiX3pFi68njg6tvy39n-U4u64HXRqu4OG7jbjZ8lWk-2EQ/exec';

declare global {
  interface Window {
    html2canvas?: (element: HTMLElement, options?: any) => Promise<HTMLCanvasElement>;
    XLSX?: any;
    ExcelJS?: any;
  }
}

const COLORS = {
  bg: '#0B0F19',
  sidebar: '#0A0E17',
  card: '#151B26',
  accent: '#00D4FF',
  text: '#F3F4F6',
  textMuted: '#9CA3AF',
  danger: '#EF4444',
  success: '#10B981',
  warning: '#F59E0B',
};

const LOGO_URL =
  'https://englishclub.my.id/wp-content/uploads/2026/05/cropped-English-Club-Gresik-Reborn-1080-x-1350-px-2.png';

const LEVELS = [
  'Kindergarten',
  'Elementary School',
  'Junior High School',
  'Senior High School',
  'University',
  'Working Professional',
];

const CLASS_MAPPING = {
  Kindergarten: ['PAUD', 'TK A', 'TK B'],
  'Elementary School': ['Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Grade 6'],
  'Junior High School': ['Grade 7', 'Grade 8', 'Grade 9'],
  'Senior High School': ['Grade 10', 'Grade 11', 'Grade 12'],
  University: ['University'],
  'Working Professional': ['Professional'],
};

const SESSIONS = [
  'PAUD - TK A - TK B Session',
  'Grade 1 - 2 Session',
  'Grade 3 - 4 Session',
  'Grade 5 - 6 Session',
  'JHS - SHS - UNI - PRO Session',
];

const getSessionGroup = (className) => {
  if (!className) return SESSIONS[4];
  if (['PAUD', 'TK A', 'TK B'].includes(className)) return SESSIONS[0];
  if (['Grade 1', 'Grade 2'].includes(className)) return SESSIONS[1];
  if (['Grade 3', 'Grade 4'].includes(className)) return SESSIONS[2];
  if (['Grade 5', 'Grade 6'].includes(className)) return SESSIONS[3];
  return SESSIONS[4];
};

// Global helper: Get student's effective session (considers override or falls back to class-based)
const getStudentSession = (student) => {
  if (!student) return SESSIONS[4];
  if (student.sessionOverride && student.sessionOverride !== 'Default') {
    return student.sessionOverride;
  }
  return getSessionGroup(student.class);
};

// Fix 4: Returns true if an attendance record's sessionGroup matches the student's CURRENT effective session.
// This prevents overridden students' old-session records from polluting payment/report calculations.
const isAttendanceValidForStudent = (attRecord, student) => {
  if (!attRecord || !student) return false;
  const currentSession = getStudentSession(student);
  // If the record has a sessionGroup, it must match the student's current effective session
  if (attRecord.sessionGroup) {
    return attRecord.sessionGroup === currentSession;
  }
  // Older records may not have sessionGroup — fall back to accepting them
  return true;
};

// Global helper: Normalisasi No. WhatsApp ke format internasional (628xxxxxxxxx)
// Menerima input lokal (08...), sudah internasional (628... / +628...), atau legacy string ber-apostrof.
// Hasilnya SELALU string digit murni siap pakai untuk link wa.me/... tanpa perlu trik apostrof lagi,
// karena format 628... tidak diawali angka 0 sehingga aman dari masalah auto-konversi Google Sheets.
const normalizeWhatsapp = (input) => {
  const digits = String(input || '').replace(/^'/, '').replace(/[^0-9]/g, '');
  if (!digits) return '';
  if (digits.startsWith('62')) return digits;
  if (digits.startsWith('0')) return '62' + digits.slice(1);
  return '62' + digits;
};

// PERBAIKAN KRITIS: Memaksa aplikasi menggunakan waktu lokal perangkat (Local Time) 
// untuk menghindari pergeseran tanggal UTC saat aplikasi diakses pagi hari (sebelum jam 7 pagi WIB).
const getTodayDateLocal = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

// Helper: Hasilkan timestamp lokal perangkat dalam format "YYYY-MM-DD HH:mm:ss"
// Menggantikan new Date().toISOString() yang selalu menghasilkan UTC (Z),
// sehingga timestamp di System Logs, Recycle Bin, dan audit trail selalu sesuai waktu lokal.
const getLocalTimestamp = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
};

// Helper: Normalisasi berbagai format timestamp (ISO UTC "Z", ISO offset "+07:00",
// atau sudah lokal "YYYY-MM-DD HH:mm:ss") menjadi string tampilan lokal yang seragam.
const normalizeTimestamp = (raw) => {
  if (!raw) return '-';
  // Sudah format lokal "YYYY-MM-DD HH:mm:ss" — langsung pakai
  if (typeof raw === 'string' && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(raw)) return raw;
  // ISO dengan timezone offset atau Z — parse lalu konversi ke waktu lokal
  const d = new Date(raw);
  if (isNaN(d.getTime())) return String(raw);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
};

// BUG FIX #9: Shared helper — merges prevDb with a normalized cloudDb.
// Previously this logic was copy-pasted 3 times (startup, login, refreshBeforeEdit).
// Now all three call this single function to stay in sync with any future changes.
const MERGE_ALL_COLS = [
  'users', 'students', 'tutors', 'studentAttendance', 'tutorAttendance',
  'journals', 'assessments', 'payments', 'payroll', 'calendar',
  'announcements', 'recycleBin', 'materials'
];

// Koleksi yang benar-benar disinkronkan. Logs TIDAK pernah ikut delta sync utama.
const SYNC_COLLECTIONS = [...MERGE_ALL_COLS];

// Timestamp untuk konflik antar-device harus timezone-neutral.
// getLocalTimestamp() tetap dipakai untuk UI/log audit, sedangkan updatedAt memakai UTC.
const getSyncTimestamp = () => new Date().toISOString();

const getSyncSnapshot = (sourceDb: any): any => {
  const snapshot: any = {};
  SYNC_COLLECTIONS.forEach(col => {
    snapshot[col] = Array.isArray(sourceDb?.[col]) ? sourceDb[col] : [];
  });
  return snapshot;
};

const getRecordId = (item: any): string => {
  if (!item || item.id === undefined || item.id === null || item.id === '') return '';
  return String(item.id);
};

// Perbandingan record harus bebas dari urutan key object. Urutan kolom di Google
// Sheets dapat berbeda dari urutan object localStorage, tetapi datanya tetap sama.
const canonicalizeForSync = (value: any): any => {
  if (Array.isArray(value)) return value.map(canonicalizeForSync);
  if (value && typeof value === 'object') {
    const out: any = {};
    Object.keys(value).sort().forEach(key => { out[key] = canonicalizeForSync(value[key]); });
    return out;
  }
  return value;
};

const syncRecordsEqual = (a: any, b: any): boolean => {
  try {
    return JSON.stringify(canonicalizeForSync(a)) === JSON.stringify(canonicalizeForSync(b));
  } catch (e) {
    return false;
  }
};

const getPendingChangesRelativeToCloud = (localDb: any, cloudDb: any, options: { includeDeletions?: boolean } = {}) => {
  const localPriorityByCollection: Record<string, Set<string>> = {};
  const deletedIdsByCollection: Record<string, Set<string>> = {};
  let hasChanges = false;
  const includeDeletions = options.includeDeletions === true;

  SYNC_COLLECTIONS.forEach(col => {
    const localArr = Array.isArray(localDb?.[col]) ? localDb[col] : [];
    const cloudArr = Array.isArray(cloudDb?.[col]) ? cloudDb[col] : [];
    const localById = new Map<string, any>();
    const cloudById = new Map<string, any>();

    localArr.forEach(item => {
      const id = getRecordId(item);
      if (id) localById.set(id, item);
    });
    cloudArr.forEach(item => {
      const id = getRecordId(item);
      if (id) cloudById.set(id, item);
    });

    const localPriority = new Set<string>();
    const localDeleted = new Set<string>();

    localArr.forEach(item => {
      const id = getRecordId(item);
      if (!id) return;
      const cloudItem = cloudById.get(id);
      if (!cloudItem) {
        // Saat pull awal, record lokal yang tidak ada di cloud dianggap
        // kandidat perubahan lokal, bukan deletion terhadap cloud.
        localPriority.add(id);
        return;
      }
      if (!syncRecordsEqual(item, cloudItem)) {
        const localTime = parseComparableTime(item);
        const cloudTime = parseComparableTime(cloudItem);
        if (localTime > cloudTime && localTime > 0) localPriority.add(id);
      }
    });

    if (includeDeletions) {
      cloudArr.forEach(item => {
        const id = getRecordId(item);
        if (id && !localById.has(id)) localDeleted.add(id);
      });
    }

    if (localPriority.size) {
      localPriorityByCollection[col] = localPriority;
      hasChanges = true;
    }
    if (localDeleted.size) {
      deletedIdsByCollection[col] = localDeleted;
      hasChanges = true;
    }
  });

  return { localPriorityByCollection, deletedIdsByCollection, hasChanges };
};

const getPendingSyncOptionsFromBaseline = (localDb: any, baselineSnapshot: string | null) => {
  if (!baselineSnapshot) {
    return { localPriorityByCollection: {}, deletedIdsByCollection: {}, hasChanges: false };
  }
  try {
    const baseline = JSON.parse(baselineSnapshot);
    const diff = buildSyncDelta(localDb, baseline);
    const localPriorityByCollection: Record<string, Set<string>> = {};
    const deletedIdsByCollection: Record<string, Set<string>> = {};

    Object.keys(diff.deltaPayload).forEach(col => {
      const ids = new Set<string>();
      (diff.deltaPayload[col] || []).forEach(item => {
        const id = getRecordId(item);
        if (id) ids.add(id);
      });
      if (ids.size) localPriorityByCollection[col] = ids;
    });

    diff.deletions.forEach(d => {
      if (!deletedIdsByCollection[d.collection]) deletedIdsByCollection[d.collection] = new Set<string>();
      deletedIdsByCollection[d.collection].add(String(d.id));
    });

    return {
      localPriorityByCollection,
      deletedIdsByCollection,
      hasChanges: diff.hasChanges
    };
  } catch (e) {
    return { localPriorityByCollection: {}, deletedIdsByCollection: {}, hasChanges: false };
  }
};

const mergeCloudData = (
  prevDb: any,
  normalizedCloud: any,
  options: {
    localPriorityByCollection?: Record<string, Set<string>>;
    deletedIdsByCollection?: Record<string, Set<string>>;
  } = {}
): any => {
  const localBin = Array.isArray(prevDb.recycleBin) ? prevDb.recycleBin : [];
  const cloudBinIdsToDelete = options.deletedIdsByCollection?.recycleBin || new Set<string>();
  const cloudBin = (Array.isArray(normalizedCloud.recycleBin) ? normalizedCloud.recycleBin : [])
    .filter((b: any) => !cloudBinIdsToDelete.has(String(b?.id || b?.binId || '')));

  const combinedBin = mergeByIds(
    localBin,
    cloudBin,
    [],
    options.localPriorityByCollection?.recycleBin || new Set<string>()
  );

  const merged: any = { ...normalizedCloud, recycleBin: combinedBin };
  MERGE_ALL_COLS.filter(col => col !== 'recycleBin').forEach(col => {
    const local = Array.isArray(prevDb[col]) ? prevDb[col] : [];
    const cloud = Array.isArray(normalizedCloud[col]) ? normalizedCloud[col] : [];
    merged[col] = mergeByIds(
      local,
      cloud,
      combinedBin,
      options.localPriorityByCollection?.[col] || new Set<string>(),
      options.deletedIdsByCollection?.[col] || new Set<string>()
    );
  });
  return merged;
};

const parseComparableTime = (item: any): number => {
  if (!item) return 0;
  const t = item.updatedAt || item.timestamp || item.lastEditedAt || item.date;
  if (!t) return 0;
  const ms = new Date(String(t).replace(' ', 'T')).getTime();
  return isNaN(ms) ? 0 : ms;
};

const buildSyncDelta = (latestDb: any, lastSynced: any | null) => {
  const deltaPayload: Record<string, any[]> = {};
  const deletions: Array<{ collection: string; id: string }> = [];
  const dbSnapshotAtRequest = getSyncSnapshot(latestDb);

  if (!lastSynced) {
    SYNC_COLLECTIONS.forEach(col => {
      const rows = Array.isArray(latestDb?.[col]) ? latestDb[col] : [];
      if (rows.length) deltaPayload[col] = rows;
    });
    return { deltaPayload, deletions, dbSnapshotAtRequest, hasChanges: Object.keys(deltaPayload).length > 0 };
  }

  SYNC_COLLECTIONS.forEach(col => {
    const localArr = Array.isArray(latestDb?.[col]) ? latestDb[col] : [];
    const baseArr = Array.isArray(lastSynced?.[col]) ? lastSynced[col] : [];
    const baseById = new Map<string, any>();
    const localIds = new Set<string>();

    baseArr.forEach(item => {
      const id = getRecordId(item);
      if (id) baseById.set(id, item);
    });
    localArr.forEach(item => {
      const id = getRecordId(item);
      if (id) localIds.add(id);
    });

    const changedRows: any[] = [];
    localArr.forEach(item => {
      const id = getRecordId(item);
      // Semua record yang tidak mempunyai ID tetap dikirim agar data legacy tidak hilang.
      if (!id) {
        changedRows.push(item);
        return;
      }
      const baseItem = baseById.get(id);
      if (!baseItem || !syncRecordsEqual(baseItem, item)) {
        changedRows.push(item);
      }
    });

    if (changedRows.length) deltaPayload[col] = changedRows;

    baseById.forEach((_baseItem, id) => {
      if (!localIds.has(id)) deletions.push({ collection: col, id });
    });
  });

  return {
    deltaPayload,
    deletions,
    dbSnapshotAtRequest,
    hasChanges: Object.keys(deltaPayload).length > 0 || deletions.length > 0
  };
};

// BUGFIX LWW, ORPHAN & PURGE: Merge dua array berdasarkan field 'id' unik dengan Last Write Wins.
const THIRTY_DAYS_MS = 30 * 24 * 60 * 60 * 1000;

const mergeByIds = (
  local: any[],
  cloud: any[],
  recycleBin: any[] = [],
  localPriorityIds: Set<string> = new Set<string>(),
  deletedIds: Set<string> = new Set<string>()
): any[] => {
  const localArr = Array.isArray(local) ? local : [];
  const cloudArr = Array.isArray(cloud) ? cloud : [];
  
  // 1. Auto-Purge: Hanya baca item tong sampah yang usianya di bawah 30 Hari
  // BUGFIX KRITIS (SISWA MUNCUL KEMBALI): Gunakan `isNaN(delTime) || ...` agar
  // bin item dengan deletedAt yang gagal di-parse (reformatted oleh GAS/Sheets)
  // TETAP masuk ke activeBin dan binSet. Logika ini konsisten dengan normalizeData.
  // Versi lama `!isNaN(delTime) && ...` membuat bin item ber-deletedAt NaN dibuang
  // dari activeBin → binSet kosong → student tidak difilter → muncul kembali di cloud.
  const activeBin = recycleBin.filter(b => {
    if (!b.deletedAt) return true;
    const delTime = new Date(String(b.deletedAt).replace(' ', 'T')).getTime(); // Safe parsing
    return isNaN(delTime) || (Date.now() - delTime < THIRTY_DAYS_MS);
  });
  const binSet = new Set(activeBin.map((b: any) => String(b.data?.id)).filter(Boolean));
  
  const merged = new Map<string, any>();

  // 2. Cascade Delete Checker (Mencegah Orphaned Records)
  const isInvalid = (item: any) => {
    if (!item || !item.id) return true;
    if (binSet.has(String(item.id))) return true;
    // Efek Domino: Jika studentId atau scheduleId item ini ada di tong sampah, item ini ikut diblokir
    if (item.studentId && binSet.has(String(item.studentId))) return true;
    if (item.scheduleId && binSet.has(String(item.scheduleId))) return true;
    return false;
  };

  // 3. Last Write Wins (LWW) Time Parser
  const parseTime = (item: any) => {
    if (!item) return 0;
    const t = item.updatedAt || item.timestamp || item.lastEditedAt || item.date;
    if (!t) return 0;
    const ms = new Date(String(t).replace(' ', 'T')).getTime();
    return isNaN(ms) ? 0 : ms;
  };

  cloudArr.forEach(item => { 
    const id = getRecordId(item);
    if (!id || deletedIds.has(id)) return;
    if (!isInvalid(item)) merged.set(id, item); 
  });

  localArr.forEach(item => { 
    if (!isInvalid(item)) {
      const existing = merged.get(String(item.id));
      if (existing) {
         const timeLocal = parseTime(item);
         const timeCloud = parseTime(existing);
         
         // BUGFIX: DEEP MERGE SUBMISSIONS PADA MATERIALS (Mencegah Tugas Saling Timpa)
         if (Array.isArray(item.submissions) || Array.isArray(existing.submissions)) {
             const mergedSubs = new Map();
             (existing.submissions || []).forEach(s => mergedSubs.set(String(s.studentId), s));
             (item.submissions || []).forEach(s => {
                 const exSub = mergedSubs.get(String(s.studentId));
                 if (exSub) {
                     // LEVEL 3 DEEP MERGE: GABUNGKAN KOMENTAR (REPLIES) GURU & SISWA
                     // FIX #10: Key harus unik — tambahkan 32-char hash dari text agar
                     // dua reply di menit yang sama dari orang yang sama tidak saling menimpa.
                     const replyKey = (r) => `${r.date}|${r.senderName}|${String(r.text || '').slice(0, 32)}`;
                     const mergedReplies = new Map();
                     (exSub.replies || []).forEach(r => mergedReplies.set(replyKey(r), r));
                     (s.replies || []).forEach(r => mergedReplies.set(replyKey(r), r));
                     
                     const timeS = parseTime(s);
                     const timeEx = parseTime(exSub);
                     
                     if (timeS > timeEx) {
                         s.replies = Array.from(mergedReplies.values());
                         mergedSubs.set(String(s.studentId), s);
                     } else {
                         exSub.replies = Array.from(mergedReplies.values());
                         mergedSubs.set(String(s.studentId), exSub);
                     }
                 } else {
                     mergedSubs.set(String(s.studentId), s);
                 }
             });
             // Simpan hasil gabungan ke item lokal
             item.submissions = Array.from(mergedSubs.values());
         }

         // Jika data Cloud lebih mutakhir/baru dari Lokal, Cloud dipertahankan (Tidak ditimpa)
         // Saat conflict, frontend mengirim hanya record yang benar-benar berubah.
         // Record tersebut dipertahankan dari sisi lokal sampai berhasil dipush ulang
         // menggunakan baseVersion server terbaru.
         if (localPriorityIds.has(String(item.id))) {
             merged.set(String(item.id), item);
             return;
         }

         if (timeCloud > timeLocal && timeCloud > 0) {
             // Pastikan hasil deep merge (tugas siswa) tetap selamat masuk ke versi Cloud
             if (item.submissions) existing.submissions = item.submissions;
             // BUGFIX SESSION OVERRIDE: Field-field "additive" yang hanya ada di lokal
             // (mis. sessionOverride, bonusExp, speakingChallengeCompletedCount) wajib
             // dipertahankan meski cloud menang LWW — cloud bisa saja versi lama yang
             // belum punya field ini, sehingga full-replace akan menghapusnya diam-diam.
             // FIX #11: PRESERVE_LOCAL_FIELDS — kondisi `existing[field] === undefined`
             // terlalu ketat. Cloud bisa punya 'Default' atau 0 (falsy) yang berarti
             // field itu belum pernah diset secara meaningful — lokal yang sudah diset
             // valid harus menang. Periksa juga nilai 'Default' dan 0 sebagai "not set".
             const PRESERVE_LOCAL_FIELDS = ['sessionOverride', 'bonusExp', 'speakingChallengeCompletedCount', 'lastSpeakingChallengeDate', 'teacherComment'];
             PRESERVE_LOCAL_FIELDS.forEach(field => {
               const localVal = item[field];
               const cloudVal = existing[field];
               const localIsSet = localVal !== undefined && localVal !== null && localVal !== 'Default';
               const cloudIsUnset = cloudVal === undefined || cloudVal === null || cloudVal === 'Default';
               if (localIsSet && cloudIsUnset) {
                 existing[field] = localVal;
               }
             });
             return;
         }
      }
      merged.set(String(item.id), item); 
    }
  });

  return Array.from(merged.values());
};

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const formatDropdownDate = (dateStr) => {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-');
  const dateObj = new Date(y, m - 1, d);
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return `${days[dateObj.getDay()]}, ${parseInt(d, 10)} ${MONTHS[parseInt(m, 10)-1]} ${y}`;
};

// BUGFIX #7: Parse "YYYY-MM-DD" secara eksplisit menggunakan new Date(y, m-1, d)
// untuk menghindari masalah UTC vs Local yang muncul saat pakai new Date("YYYY-MM-DD").
// "YYYY-MM-DD" diparse sebagai UTC midnight oleh browser, sehingga di timezone
// dengan offset negatif bisa muncul 1 hari lebih awal.
const safeDateDisplay = (dateStr: string, locale: string, options: Intl.DateTimeFormatOptions) => {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return d.toLocaleDateString(locale, options);
};

const calculateDaysLeft = (targetDate, todayDate) => {
  // PERBAIKAN KRITIS: Parsing string secara eksplisit agar bebas dari pengaruh Timezone Browser
  if (!targetDate || !todayDate) return '-';
  const [tYear, tMonth, tDay] = targetDate.split('-').map(Number);
  const [dYear, dMonth, dDay] = todayDate.split('-').map(Number);
  
  const t = new Date(tYear, tMonth - 1, tDay).getTime();
  const d = new Date(dYear, dMonth - 1, dDay).getTime();
  const diff = Math.round((t - d) / (1000 * 3600 * 24));
  
  if (diff === 0) return 'Today';
  if (diff === 1) return 'Tomorrow';
  if (diff < 0) return 'Passed';
  return `${diff} Days Left`;
};

const parseSessions = (sessionStr) => {
  if (!sessionStr) return [];
  if (Array.isArray(sessionStr)) return sessionStr;
  return String(sessionStr).split('|').map(s => s.trim()).filter(Boolean);
};

const getPerformanceCat = (score) => {
  if (score >= 90) return { label: 'Excellent', color: 'text-green-600', bg: 'bg-green-100' };
  if (score >= 80) return { label: 'Very Good', color: 'text-blue-600', bg: 'bg-blue-100' };
  if (score >= 70) return { label: 'Good', color: 'text-teal-600', bg: 'bg-teal-100' };
  if (score >= 60) return { label: 'Fair', color: 'text-yellow-600', bg: 'bg-yellow-100' };
  return { label: 'Needs Improvement', color: 'text-red-600', bg: 'bg-red-100' };
};

// --- NEW HELPERS FOR GAMIFICATION ---
const calculateStudentEXP = (studentId, db) => {
   const expFromAtt = db.studentAttendance.filter(a => a.studentId === studentId && a.status === 'Present').length * 50;
   const expFromScore = db.assessments.filter(a => a.studentId === studentId && a.average >= 80).length * 100;
   let totalTasksDone = 0;
   (db.materials || []).forEach(m => {
      if ((m.submissions || []).some(s => s.studentId === studentId)) totalTasksDone++;
   });
   
   // Tambahan EXP dari Daily Speaking Challenge
   const student = db.students.find(s => s.id === studentId);
   const expFromSpeaking = (student?.speakingChallengeCompletedCount || 0) * 20; // 20 EXP per challenge
   
   const bonusExp = Number(student?.bonusExp) || 0; // EXP manual dari tutor/admin
   
   return expFromAtt + expFromScore + (totalTasksDone * 30) + expFromSpeaking + bonusExp;
};

const getLevelInfo = (exp) => {
   if (exp >= 2000) return { title: 'Diamond Scholar', color: 'from-cyan-300 to-blue-500', bg: 'bg-cyan-500/10', border: 'border-cyan-500/50', icon: Trophy, textCol: 'text-cyan-400' };
   if (exp >= 1000) return { title: 'Gold Achiever', color: 'from-amber-300 to-orange-500', bg: 'bg-amber-500/10', border: 'border-amber-500/50', icon: Medal, textCol: 'text-amber-400' };
   if (exp >= 400) return { title: 'Silver Explorer', color: 'from-slate-300 to-gray-400', bg: 'bg-slate-500/10', border: 'border-slate-500/50', icon: ShieldCheck, textCol: 'text-slate-300' };
   return { title: 'Bronze Beginner', color: 'from-orange-700 to-amber-900', bg: 'bg-orange-900/20', border: 'border-orange-900/50', icon: Star, textCol: 'text-orange-500' };
};

const getLeaderboardBadge = (idx) => {
   if (idx === 0) return { icon: Trophy, title: '1st Gold', color: 'text-yellow-400', bg: 'bg-yellow-500/10' };
   if (idx === 1) return { icon: Medal, title: '2nd Silver', color: 'text-slate-300', bg: 'bg-slate-400/10' };
   if (idx === 2) return { icon: Medal, title: '3rd Bronze', color: 'text-amber-600', bg: 'bg-amber-600/10' };
   if (idx === 3) return { icon: Award, title: 'High Achiever', color: 'text-indigo-400', bg: 'bg-indigo-500/10' };
   if (idx === 4) return { icon: Star, title: 'Rising Star', color: 'text-pink-400', bg: 'bg-pink-500/10' };
   if (idx === 5) return { icon: Target, title: 'Sharp Learner', color: 'text-rose-400', bg: 'bg-rose-500/10' };
   if (idx === 6) return { icon: Zap, title: 'Fast Thinker', color: 'text-orange-400', bg: 'bg-orange-500/10' };
   if (idx === 7) return { icon: ShieldCheck, title: 'Steady Scholar', color: 'text-cyan-400', bg: 'bg-cyan-500/10' };
   if (idx === 8) return { icon: GraduationCap, title: 'Bright Mind', color: 'text-teal-400', bg: 'bg-teal-500/10' };
   if (idx === 9) return { icon: BookOpen, title: 'Knowledge Seeker', color: 'text-blue-400', bg: 'bg-blue-500/10' };
   return { icon: Hash, title: `Rank ${idx + 1}`, color: 'text-gray-500', bg: 'bg-gray-800/50' };
};

// --- NEW HELPER: SORT STUDENTS LOGICALLY ---
const sortStudentsLogically = (students) => {
  return [...students].sort((a, b) => {
    // 1. Sort by Level (Lowest first based on LEVELS array)
    const levelA = LEVELS.indexOf(a.level);
    const levelB = LEVELS.indexOf(b.level);
    const validLevelA = levelA !== -1 ? levelA : 999;
    const validLevelB = levelB !== -1 ? levelB : 999;
    if (validLevelA !== validLevelB) return validLevelA - validLevelB;
    
    // 2. Sort by Class (Lowest first based on CLASS_MAPPING array)
    const classA = CLASS_MAPPING[a.level]?.indexOf(a.class) ?? 999;
    const classB = CLASS_MAPPING[b.level]?.indexOf(b.class) ?? 999;
    if (classA !== classB) return classA - classB;

    // 3. Sort by Alphabet (A-Z)
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
};
// ------------------------------------

// AUTO-GENERATE TEACHER COMMENTS ALGORITHM
const generateAutoComment = (student, attRate, avgScore, assessments) => {
  if (!student) return "No student data available.";
  const sClass = student.class || '';
  const sLevel = student.level || '';
  const isKids = ['Kindergarten', 'PAUD', 'TK A', 'TK B'].includes(sLevel) || sClass.includes('TK') || sClass.includes('PAUD');
  const name = (student.name || 'Student').split(' ')[0];
  let latestAss = assessments.length > 0 ? assessments[assessments.length - 1] : null;

  let highSubj = '';
  let lowSubj = '';
  if (latestAss && latestAss.scores) {
       const scores = Object.entries(latestAss.scores).filter(([k]) => k !== 'material' && latestAss.scores[k] !== '' && latestAss.scores[k] !== undefined && !isNaN(Number(latestAss.scores[k])));
       if(scores.length > 0) {
           scores.sort((a,b) => Number(b[1]) - Number(a[1]));
           highSubj = scores[0][0];
           lowSubj = scores[scores.length - 1][0];
       }
  }

  const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];
  let comment = "";

  if (isKids) {
       // BUGFIX: avgScore === 0 berarti belum ada data nilai, jangan masuk branch "Excellent"
       if (avgScore === 0) {
           comment = `${name} baru saja bergabung dan sedang dalam proses adaptasi. Kami akan terus memantau perkembangannya dengan penuh semangat!`;
       } else if (avgScore >= 85) {
           const intros = [
               `${name} sangat ceria dan antusias dalam mengikuti kegiatan kelas. `,
               `${name} menunjukkan semangat belajar yang luar biasa dan selalu aktif berpartisipasi. `,
               `Kami sangat senang melihat perkembangan ${name} yang semakin berani dan percaya diri. `
           ];
           const strengths = highSubj ? [
               `Kemampuan ${highSubj}-nya patut diacungi jempol. `,
               `Perkembangannya dalam ${highSubj} terlihat sangat pesat. `
           ] : [`Perkembangan bahasa Inggrisnya menunjukkan kemajuan yang sangat positif. `];
           const closings = [
               `Terus pertahankan semangat belajar yang luar biasa ini ya!`,
               `Tetap semangat bermain dan belajar bahasa Inggris bersama teman-teman!`,
               `Good job!, ${name}! Teruslah berprestasi.`
           ];
           comment = pick(intros) + pick(strengths) + pick(closings);
       } else if (avgScore >= 70) {
           const intros = [
               `${name} menunjukkan ketertarikan yang baik selama sesi belajar bersama teman-teman. `,
               `Perkembangan ${name} di kelas cukup baik dan semakin terbiasa dengan lingkungan belajar. `
           ];
           const strengths = highSubj ? `Pemahaman ${highSubj} sudah mulai terlihat dengan baik. ` : '';
           const improvements = (lowSubj && lowSubj !== highSubj) ? `Dengan sedikit bimbingan lebih pada ${lowSubj}, hasilnya pasti akan lebih maksimal. ` : `Mari kita terus dorong keberaniannya dalam berekspresi. `;
           const closings = [`Tetap semangat berlatih di rumah ya!`, `Mari kita terus berikan motivasi agar ${name} semakin hebat.`];
           comment = pick(intros) + strengths + improvements + pick(closings);
       } else {
           const intros = [
               `${name} adalah anak yang hebat dan sedang terus belajar beradaptasi dengan baik. `,
               `Kami melihat potensi yang baik pada diri ${name} selama mengikuti kegiatan. `
           ];
           const improvements = lowSubj ? `Mari kita bersama-sama memberi dukungan lebih agar ${name} semakin percaya diri, khususnya dalam pengenalan ${lowSubj}. ` : `Bimbingan lebih lanjut akan sangat membantu pemahamannya. `;
           const closings = [`Semangat terus belajarnya ya, ${name}!`, `Kami yakin ${name} bisa jauh lebih baik lagi.`];
           comment = pick(intros) + improvements + pick(closings);
       }
  } else {
       // BUGFIX: avgScore === 0 berarti belum ada data nilai, jangan masuk branch "Excellent"
       if (avgScore === 0) {
           comment = `${name} baru bergabung dan kami sangat antusias menyambut kehadiran beliau. Data penilaian akan tersedia setelah mengikuti sesi evaluasi pertama.`;
       } else if (avgScore >= 85) {
           const intros = [
               `Proses belajar ${name} menunjukkan hasil yang sangat memuaskan, terlihat dari partisipasi aktif di kelas. `,
               `Kami sangat bangga dengan pencapaian akademik ${name} sejauh ini yang terus konsisten. `,
               `Selama periode ini, ${name} memperlihatkan dedikasi belajar yang luar biasa. `
           ];
           const strengths = highSubj ? [
               `Penguasaan pada aspek ${highSubj} sangat menonjol. `,
               `Kemampuan analitis dan pemahaman ${highSubj} berkembang pesat. `
           ] : [`Pemahaman materi secara komprehensif sangat patut diapresiasi. `];
           const closings = [
               `Terus pertahankan prestasi dan dedikasi cemerlang ini untuk tantangan akademik ke depannya.`,
               `Kami berharap ${name} dapat terus mempertahankan momentum positif ini.`,
               `Keep up the excellent work!`
           ];
           comment = pick(intros) + pick(strengths) + pick(closings);
       } else if (avgScore >= 70) {
           const intros = [
               `${name} memiliki pemahaman dasar yang baik dan cukup konsisten dalam mengikuti kelas. `,
               `Performa akademik ${name} menunjukkan progress yang stabil selama periode pembelajaran ini. `
           ];
           const strengths = highSubj ? `Keterampilan ${highSubj} merupakan salah satu kekuatan utamanya. ` : '';
           const improvements = lowSubj ? `Fokus tambahan pada area ${lowSubj} sangat disarankan untuk hasil yang lebih terpadu. ` : `Terdapat ruang untuk peningkatan yang lebih baik dalam kelancaran komunikasi. `;
           const closings = [
               `Terus tingkatkan intensitas latihan agar pencapaian akademik semakin maksimal.`,
               `Jangan ragu untuk terus berlatih dan bertanya. Semangat!`
           ];
           comment = pick(intros) + strengths + improvements + pick(closings);
       } else {
           const intros = [
               `${name} telah berusaha dengan baik untuk mengikuti ritme pembelajaran di kelas bahasa Inggris ini. `,
               `Kami menghargai setiap usaha yang ditunjukkan oleh ${name} selama mengikuti sesi kelas. `
           ];
           const improvements = lowSubj ? `Diperlukan dedikasi lebih dan fokus tambahan, khususnya pada keterampilan ${lowSubj}, agar dapat mengejar target pembelajaran. ` : `Dibutuhkan konsistensi ekstra untuk memahami konsep dasar yang diberikan. `;
           const closings = [
               `Jangan ragu untuk berdiskusi lebih aktif dengan tutor. Kami siap membantu!`,
               `Tingkatkan terus semangat belajarnya untuk mengejar ketertinggalan.`
           ];
           comment = pick(intros) + improvements + pick(closings);
       }
  }

  if (attRate < 75 && attRate > 0) {
      const attWarnings = [
          ` Selain itu, kami berharap tingkat kehadiran ${name} dapat lebih ditingkatkan agar tidak tertinggal materi penting.`,
          ` Peningkatan kehadiran di kelas juga akan sangat berdampak positif bagi progress belajarnya.`
      ];
      comment += pick(attWarnings);
  }

  return comment;
};

const Card = ({ children, className = '', id, onClick }: any) => (
  <div
    id={id}
    onClick={onClick}
    className={`bg-[#151B26] border border-gray-800 rounded-xl p-4 sm:p-5 shadow-xl ${className}`}
  >
    {children}
  </div>
);

// GlassCard: consistent glassmorphism style used in Journals, Assessments, Announcements
const GlassCard = ({ children, className = '', accentColor = 'cyan', onClick }: any) => {
  const accents: Record<string, string> = {
    cyan:   'border-[#00D4FF]/20 shadow-[0_4px_24px_rgba(0,212,255,0.06)]',
    yellow: 'border-yellow-400/20 shadow-[0_4px_24px_rgba(234,179,8,0.06)]',
    purple: 'border-purple-500/20 shadow-[0_4px_24px_rgba(168,85,247,0.06)]',
    emerald:'border-emerald-500/20 shadow-[0_4px_24px_rgba(16,185,129,0.06)]',
    amber:  'border-amber-500/20 shadow-[0_4px_24px_rgba(245,158,11,0.06)]',
  };
  return (
    <div
      onClick={onClick}
      className={`bg-[#151B26]/90 backdrop-blur-sm border rounded-xl p-4 sm:p-5 ${accents[accentColor] || accents.cyan} ${className}`}
    >
      {children}
    </div>
  );
};

const Button = ({
  children,
  onClick,
  type = 'button',
  variant = 'primary',
  className = '',
  icon: Icon,
  disabled = false,
}: any) => {
  const baseStyle =
    'flex items-center justify-center gap-2 px-4 py-2 rounded-lg font-medium transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed';
  const variants = {
    primary:
      'bg-[#00D4FF] text-[#0B0F19] hover:bg-[#00b8e6] shadow-[0_0_15px_rgba(0,212,255,0.3)]',
    secondary:
      'bg-gray-800 text-white hover:bg-gray-700 border border-gray-700',
    danger:
      'bg-red-500/10 text-red-500 hover:bg-red-500/20 border border-red-500/20',
    ghost: 'hover:bg-gray-800 text-gray-300 hover:text-white',
  };
  return (
    <button
      type={type as 'button'}
      onClick={onClick}
      disabled={disabled}
      className={`${baseStyle} ${variants[variant]} ${className}`}
    >
      {Icon && <Icon size={18} />}
      {children}
    </button>
  );
};

const Input = ({
  label,
  type = 'text',
  value,
  onChange,
  options = [],
  disabled = false,
  required = false,
  placeholder = '',
  className = '',
  min,
  max,
}: any) => {
  const baseClass =
    'w-full bg-[#0B0F19] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] focus:ring-1 focus:ring-[#00D4FF] transition-all disabled:opacity-50';
  return (
    <div className={`w-full ${className}`}>
      {label && (
        <label className="block text-sm text-gray-400 mb-1.5">
          {label} {required && <span className="text-red-500">*</span>}
        </label>
      )}
      {type === 'select' ? (
        <select
          value={value}
          onChange={(e) => onChange(e.target.value)}
          disabled={disabled}
          className={baseClass}
          required={required}
        >
          <option value="">Select...</option>
          {options.map((opt, i) => (
            <option key={i} value={opt.value || opt}>
              {opt.label || opt}
            </option>
          ))}
        </select>
      ) : (
        <input
          type={type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          disabled={disabled}
          className={baseClass}
          required={required}
          placeholder={placeholder}
          min={min}
          max={max}
        />
      )}
    </div>
  );
};

// Standardized empty-state used across all tables/lists (icon + title + description)
const EmptyState = ({ icon: Icon = Inbox, title, description, className = '', animated = false }) => (
  <div className={`flex flex-col items-center justify-center py-10 px-4 text-center ${className}`}>
    <div className={`relative mb-4 ${animated ? 'animate-pulse' : ''}`}>
      {/* Decorative glow ring */}
      <div className="absolute inset-0 rounded-full bg-[#00D4FF]/5 blur-xl scale-150 pointer-events-none" />
      <div className="relative w-16 h-16 rounded-2xl bg-[#0B0F19] border border-gray-800 flex items-center justify-center shadow-xl">
        <Icon size={28} className="text-gray-600" />
      </div>
    </div>
    <p className="font-bold text-base sm:text-lg text-white mb-1">{title}</p>
    {description && <p className="text-sm text-gray-500 max-w-sm leading-relaxed">{description}</p>}
  </div>
);

// Alias for consistent status badge usage across all modules
const StatusBadge = ({ status }: { status: string }) => <Badge status={status} />;

const Badge = ({ status }) => {
  const styles = {
    Active: 'bg-green-500/10 text-green-400 border-green-500/20',
    Inactive: 'bg-red-500/10 text-red-400 border-red-500/20',
    Present: 'bg-green-500/10 text-green-400 border-green-500/20',
    Sick: 'bg-yellow-500/10 text-yellow-400 border-yellow-500/20',
    Excused: 'bg-purple-500/10 text-purple-400 border-purple-500/20',
    Absent: 'bg-red-500/10 text-red-400 border-red-500/20',
    Paid: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/50 shadow-[0_0_10px_rgba(16,185,129,0.4)]',
    Unpaid: 'bg-rose-500/20 text-rose-400 border-rose-500/50 shadow-[0_0_10px_rgba(244,63,94,0.4)]',
    Partial: 'bg-amber-500/20 text-amber-400 border-amber-500/50 shadow-[0_0_10px_rgba(245,158,11,0.4)]',
    'No Target': 'bg-gray-500/10 text-gray-400 border-gray-500/20',
    Draft: 'bg-gray-500/10 text-gray-400 border-gray-500/20',
    Debt: 'bg-red-500/20 text-red-400 border-red-500/50 shadow-[0_0_10px_rgba(244,63,94,0.4)]',
    Deposit: 'bg-cyan-500/20 text-cyan-400 border-cyan-500/50 shadow-[0_0_10px_rgba(6,182,212,0.4)]',
  };
  
  const isProminent = ['Paid', 'Unpaid', 'Partial', 'Debt', 'Deposit'].includes(status);

  return (
    <span
      className={`px-2.5 py-1 text-xs rounded-full border whitespace-nowrap ${
        isProminent ? 'font-black uppercase tracking-wider' : 'font-medium'
      } ${styles[status] || 'bg-gray-800 text-gray-300'}`}
    >
      {status}
    </span>
  );
};

const NewBadge = ({ isNew, billingStartMonth = null }: { isNew?: string, billingStartMonth?: string | null }) => {
  // Helper: cek apakah billing belum dimulai di bulan berjalan
  const isBillingPending = (() => {
    if (!billingStartMonth) return false;
    const [bYear, bMonth] = billingStartMonth.split('-').map(Number);
    if (!bYear || !bMonth) return false;
    const now = new Date();
    const nowYear = now.getFullYear();
    const nowMonth = now.getMonth() + 1;
    // Pending jika bulan berjalan masih sebelum billingStartMonth
    return (nowYear < bYear) || (nowYear === bYear && nowMonth < bMonth);
  })();

  return (
    <>
      {isNew === 'New' && (
        <span className="ml-2 px-1.5 py-0.5 bg-pink-500/20 text-pink-400 border border-pink-500/30 rounded text-[11px] font-bold uppercase tracking-widest shadow-[0_0_10px_rgba(236,72,153,0.2)] align-middle shrink-0">
          NEW
        </span>
      )}
      {isBillingPending && (
        <span
          title={`Mulai tagih: ${billingStartMonth}`}
          className="ml-2 px-1.5 py-0.5 bg-amber-500/20 text-amber-400 border border-amber-500/30 rounded text-[11px] font-bold uppercase tracking-widest shadow-[0_0_10px_rgba(245,158,11,0.2)] align-middle shrink-0 whitespace-nowrap"
        >
          PENDING BILLING
        </span>
      )}
    </>
  );
};

const CustomModal = ({ isOpen, onClose, title, children, zIndexClass = 'z-50' }) => {
  if (!isOpen) return null;
  return (
    <div className={`fixed inset-0 ${zIndexClass} flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animation-fade-in print:hidden`}>
      <div className="bg-[#151B26] border border-gray-700 rounded-xl w-full max-w-lg shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        <div className="p-4 border-b border-gray-800 flex justify-between items-center bg-[#0A0E17]">
          <h3 className="text-lg font-bold text-white">{title}</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-white">
            <X size={20} />
          </button>
        </div>
        <div className="p-5 pb-8 overflow-y-auto">{children}</div>
      </div>
    </div>
  );
};

const generateDummyDatabase = () => {
  const db = {
    users: [
      {
        id: 'ADM-001',
        username: 'admin',
        password: 'password',
        role: 'admin',
        name: 'Admin ECG',
        active: 'Active',
      },
    ],
    students: [],
    tutors: [],
    studentAttendance: [],
    tutorAttendance: [],
    journals: [],
    assessments: [],
    payments: [],
    payroll: [],
    calendar: [],
    announcements: [],
    recycleBin: [],
    materials: [],
  };

  return db;
};

// CUSTOM HOOK: Debounce untuk performa pencarian (Search)
// Hook: Persist a value to localStorage so filters survive module navigation
function useLocalStorage<T>(key: string, defaultValue: T): [T, (v: T) => void] {
  const [value, setValue] = useState<T>(() => {
    try {
      const stored = localStorage.getItem(key);
      return stored !== null ? JSON.parse(stored) : defaultValue;
    } catch {
      return defaultValue;
    }
  });
  const set = (v: T) => {
    setValue(v);
    try { localStorage.setItem(key, JSON.stringify(v)); } catch {}
  };
  return [value, set];
}

function useDebounce<T>(value: T, delay: number): T {
  const [debouncedValue, setDebouncedValue] = useState<T>(value);
  useEffect(() => {
    const handler = setTimeout(() => setDebouncedValue(value), delay);
    return () => clearTimeout(handler);
  }, [value, delay]);
  return debouncedValue;
}

// COMPONENT UNTUK FORCE PASSWORD CHANGE SISWA PADA LOGIN PERTAMA
function ForcePasswordChangeScreen({ user, db, setDb, setCurrentUser, showToast, setSidebarOpen }) {
  const [newPwd, setNewPwd] = useState('');
  const [confirmPwd, setConfirmPwd] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    if (newPwd !== confirmPwd) return showToast('Passwords do not match!', 'error');
    if (newPwd.length < 6) return showToast('Password must be at least 6 characters.', 'warning');

    if (user.role === 'tutor') {
      setDb(prev => ({
        ...prev,
        tutors: prev.tutors.map(t => t.username === user.username ? { ...t, password: newPwd, mustChangePassword: false } : t)
      }));
    } else {
      setDb(prev => ({
        ...prev,
        users: prev.users.map(u => u.username === user.username ? { ...u, password: newPwd, mustChangePassword: false } : u)
      }));
    }

    // PERBAIKAN: Perbarui Session Storage agar user tidak diminta ganti sandi lagi saat me-refresh halaman
    setCurrentUser(prev => {
      const updatedUser = { ...prev, password: newPwd, mustChangePassword: false };
      sessionStorage.setItem('ecg_active_session', JSON.stringify(updatedUser));
      return updatedUser;
    });
    
    showToast('Password updated successfully! Welcome to Academic Suite.', 'success');
    if (setSidebarOpen) setSidebarOpen(true);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-[#0B0F19] p-4 font-sans text-white">
      <Card className="w-full max-w-md border border-[#00D4FF]/20 shadow-[0_0_30px_rgba(0,212,255,0.1)]">
        <h2 className="text-2xl font-bold mb-2 text-[#00D4FF]">Set New Password</h2>
        <p className="text-gray-400 mb-6 text-sm">For your security, please change your temporary password before continuing to Academic Suite.</p>
        <form onSubmit={handleSubmit} className="space-y-4">
          <Input label="New Password" type="password" value={newPwd} onChange={setNewPwd} required placeholder="Min. 6 characters" />
          <Input label="Confirm New Password" type="password" value={confirmPwd} onChange={setConfirmPwd} required placeholder="Re-type new password" />
          <Button type="submit" className="w-full mt-4 py-3">Save and Continue</Button>
        </form>
      </Card>
    </div>
  );
}

function LoginScreen({ onLogin, isDbLoaded = true, language = 'en', setLanguage }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [loginStatus, setLoginStatus] = useState(''); // progressive status messages
  const [loginError, setLoginError] = useState('');
  const [rememberMe, setRememberMe] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem('ecg_remembered_user');
    if (saved) {
      try {
        // SECURITY FIX: Hanya isi username, password TIDAK disimpan di localStorage
        const { username: savedUsername } = JSON.parse(saved);
        if (savedUsername) {
          setUsername(savedUsername);
          setRememberMe(true);
        }
      } catch (e) {}
    }
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    setLoginError('');
    setLoginStatus(language === 'id' ? 'Menghubungi server...' : 'Connecting to server...');

    // After 2s, update status so user knows it's still working (GAS cold start)
    const t1 = setTimeout(() => setLoginStatus(language === 'id' ? 'Server sedang bersiap...' : 'Server is warming up...'), 2000);
    // After 5s, reassure user we're still trying
    const t2 = setTimeout(() => setLoginStatus(language === 'id' ? 'Hampir selesai, harap tunggu...' : 'Almost there, please wait...'), 5000);

    const result = await onLogin(username, password, rememberMe);

    clearTimeout(t1);
    clearTimeout(t2);
    setLoginStatus('');
    if (!result || !result.success) {
      setLoginError(result?.error || (language === 'id' ? 'Nama pengguna atau kata sandi salah. Silakan coba lagi.' : 'Invalid username or password. Please try again.'));
    }
    setIsLoading(false);
  };

  return (
    <div className="min-h-screen w-full flex bg-[#051126] font-['Poppins',sans-serif] selection:bg-[#00C2FF] selection:text-[#051126] overflow-hidden relative">
      
      {/* Language Switcher Absolute */}
      <div className="absolute top-4 right-4 sm:top-6 sm:right-6 z-50 flex items-center gap-2 bg-[#151B26]/80 backdrop-blur-md p-1.5 rounded-xl border border-gray-800 shadow-lg">
         <span className="text-xs text-gray-400 font-semibold px-2 hidden sm:block">
           {language === 'id' ? 'Bahasa Pilihan' : 'Preferred Language'}
         </span>
         <div className="flex bg-[#0B0F19] rounded-md border border-gray-700 overflow-hidden">
            <button type="button" onClick={() => setLanguage && setLanguage('en')} className={`px-3 py-1.5 text-[11px] font-bold transition-colors ${language === 'en' ? 'bg-[#00C2FF] text-[#051126]' : 'text-gray-500 hover:text-white'}`}>EN</button>
            <button type="button" onClick={() => setLanguage && setLanguage('id')} className={`px-3 py-1.5 text-[11px] font-bold transition-colors ${language === 'id' ? 'bg-[#00C2FF] text-[#051126]' : 'text-gray-500 hover:text-white'}`}>ID</button>
         </div>
      </div>

      {/* LEFT SIDE - 60% PREMIUM BRANDING */}
      <div className="hidden lg:flex lg:w-[60%] relative flex-col justify-center p-12 xl:p-24 overflow-hidden bg-gradient-to-br from-[#0A3D91] to-[#051126]">
         <div className="absolute top-[-10%] left-[-10%] w-[50%] h-[50%] bg-[#00C2FF]/20 blur-[120px] rounded-full pointer-events-none"></div>
         <div className="absolute bottom-[-20%] right-[-10%] w-[60%] h-[60%] bg-[#00C2FF]/10 blur-[150px] rounded-full pointer-events-none"></div>

         <div className="relative z-10 max-w-3xl animation-fade-in">
           <div className="flex items-center gap-6 mb-8">
             <img src={LOGO_URL} alt="ECG Logo" className="h-20 xl:h-24 w-auto drop-shadow-[0_0_20px_rgba(0,194,255,0.4)]" />
             <div className="border-l-2 border-white/20 pl-6">
               <h1 className="text-3xl xl:text-[2.5rem] font-bold text-white tracking-tight leading-tight mb-1">English Club Gresik</h1>
               <h2 className="text-xl xl:text-2xl font-extrabold text-[#00C2FF] tracking-widest uppercase">Academic Suite</h2>
             </div>
           </div>
           
           <h3 className="text-2xl text-white mb-4 font-semibold tracking-tight">Educational Administration Platform</h3>
           <p className="text-base text-blue-100/70 mb-10 leading-relaxed max-w-xl">
             Manage students, tutors, attendance, assessments, learning journals, payments, reports, and academic activities through one integrated platform designed for modern education.
           </p>

           <div className="bg-white/[0.03] backdrop-blur-xl border border-white/10 p-8 rounded-[24px] shadow-2xl max-w-2xl">
             <p className="text-xs text-[#00C2FF] font-black uppercase tracking-widest mb-6">Platform Capabilities</p>
             <div className="grid grid-cols-2 gap-x-8 gap-y-5">
               {[
                 'Student Management', 'Tutor Management', 'Attendance Tracking',
                 'Monthly Assessments', 'Learning Journals', 'Financial Management',
                 'Reports & Analytics', 'Academic Calendar'
               ].map((feature, i) => (
                 <div key={i} className="flex items-center gap-3 group">
                   <div className="flex-shrink-0 w-6 h-6 rounded-full bg-[#00C2FF]/10 flex items-center justify-center border border-[#00C2FF]/30 group-hover:bg-[#00C2FF]/20 group-hover:scale-110 transition-all duration-300 shadow-[0_0_10px_rgba(0,194,255,0.1)]">
                     <CheckCircle2 size={12} className="text-[#00C2FF]" />
                   </div>
                   <span className="text-sm text-blue-100/80 font-medium group-hover:text-white transition-colors">{feature}</span>
                 </div>
               ))}
             </div>
           </div>
         </div>
      </div>

      {/* RIGHT SIDE - 40% LOGIN FORM */}
      <div className="w-full lg:w-[40%] relative flex flex-col justify-center items-center p-6 sm:p-12 z-20 bg-[#0B0F19]">
        <div className="lg:hidden flex flex-col items-center mb-8 text-center animation-fade-in w-full">
           <img src={LOGO_URL} alt="ECG Logo" className="h-16 w-auto mb-4 drop-shadow-[0_0_15px_rgba(0,194,255,0.4)]" />
           <h1 className="text-2xl font-bold text-white tracking-tight">English Club Gresik</h1>
           <h2 className="text-lg font-extrabold text-[#00C2FF] tracking-widest uppercase mt-1">Academic Suite</h2>
        </div>

        <div className="w-full max-w-[440px] bg-[#151B26]/80 backdrop-blur-2xl border border-gray-800 p-8 sm:p-10 rounded-[24px] shadow-[0_8px_32px_rgba(0,0,0,0.4)] animation-fade-in" style={{ animationDelay: '100ms' }}>
           <div className="mb-8">
             <h2 className="text-3xl font-bold text-white mb-2 tracking-tight">{language === 'id' ? 'Selamat Datang Kembali' : 'Welcome Back'}</h2>
             <p className="text-gray-400 text-sm font-medium">{language === 'id' ? 'Masuk untuk melanjutkan ke dasbor Anda' : 'Sign in to continue to your dashboard'}</p>
           </div>

           <form onSubmit={handleSubmit} className="space-y-5">
             {loginError && (
               <div className="bg-red-500/10 border border-red-500/20 text-red-500 text-sm p-3.5 rounded-xl flex items-start gap-2 animation-fade-in shadow-sm">
                 <AlertCircle size={18} className="flex-shrink-0 mt-0.5" />
                 <span className="font-medium">{loginError}</span>
               </div>
             )}
             
             <div>
               <label className="block text-sm text-gray-400 mb-2 font-medium">{language === 'id' ? 'Nama Pengguna / Email' : 'Username / Email'}</label>
               <input type="text" value={username} onChange={(e) => setUsername(e.target.value)} required className="w-full bg-[#0B0F19] border border-gray-700 rounded-xl px-4 py-3.5 text-white focus:outline-none focus:border-[#00C2FF] focus:ring-1 focus:ring-[#00C2FF] transition-all shadow-inner placeholder-gray-600" placeholder={language === 'id' ? 'Masukkan nama pengguna' : 'Enter your username'} />
             </div>

             <div>
               <label className="block text-sm text-gray-400 mb-2 font-medium">{language === 'id' ? 'Kata Sandi' : 'Password'}</label>
               <div className="relative flex items-center">
                 <input type={showPassword ? "text" : "password"} value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full bg-[#0B0F19] border border-gray-700 rounded-xl pl-4 pr-12 py-3.5 text-white focus:outline-none focus:border-[#00C2FF] focus:ring-1 focus:ring-[#00C2FF] transition-all shadow-inner placeholder-gray-600" placeholder="••••••••" />
                 <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute right-4 text-gray-500 hover:text-[#00C2FF] transition-colors focus:outline-none p-1">
                    <Eye size={20} className={showPassword ? "text-[#00C2FF]" : "opacity-60"} />
                 </button>
               </div>
             </div>

             <div className="flex items-center justify-between mt-2 pt-1">
               <label className="flex items-center gap-2 cursor-pointer group">
                 <input type="checkbox" checked={rememberMe} onChange={(e) => setRememberMe(e.target.checked)} className="w-4 h-4 rounded border-gray-700 text-[#00C2FF] focus:ring-[#00C2FF] bg-[#0B0F19] cursor-pointer" />
                 <span className="text-sm text-gray-400 group-hover:text-white transition-colors font-medium">{language === 'id' ? 'Ingat Saya' : 'Remember Me'}</span>
               </label>
               <a href="https://wa.link/uwdlzm" target="_blank" rel="noopener noreferrer" className="text-sm text-[#00C2FF] hover:text-[#00A3D9] transition-colors font-semibold">{language === 'id' ? 'Lupa Kata Sandi?' : 'Forgot Password?'}</a>
             </div>

             <button type="submit" disabled={isLoading || !isDbLoaded} className="w-full mt-8 bg-[#00C2FF] hover:bg-[#00A3D9] text-[#051126] font-bold py-3.5 px-4 rounded-xl transition-all duration-300 transform hover:-translate-y-0.5 hover:shadow-[0_8px_25px_rgba(0,194,255,0.25)] disabled:opacity-70 disabled:cursor-not-allowed disabled:transform-none flex items-center justify-center gap-3 text-base">
               {!isDbLoaded ? (
                 <span className="flex items-center gap-2 font-mono tracking-widest font-bold">
                   {language === 'id' ? 'MENYIAPKAN...' : 'FINALIZING...'}
                 </span>
               ) : isLoading ? (
                 <span className="flex flex-col items-center gap-1">
                   <span className="flex items-center gap-2">
                     <span className="w-5 h-5 border-2 border-[#051126]/20 border-t-[#051126] rounded-full animate-spin inline-block"></span>
                     {language === 'id' ? 'Mengautentikasi...' : 'Authenticating...'}
                   </span>
                   {loginStatus && (
                     <span className="text-[10px] font-normal text-[#051126]/70 tracking-wide">{loginStatus}</span>
                   )}
                 </span>
               ) : (
                 language === 'id' ? 'Masuk' : 'Sign In'
               )}
             </button>
           </form>
        </div>

        <div className="mt-12 text-center space-y-3 animation-fade-in" style={{ animationDelay: '200ms' }}>
           <p className="text-sm text-gray-500 font-medium">
             {language === 'id' ? 'Butuh bantuan?' : 'Need help?'} <a href="https://wa.link/9awys1" target="_blank" rel="noopener noreferrer" className="text-[#00C2FF] hover:underline font-semibold">{language === 'id' ? 'Hubungi Dukungan' : 'Contact Support'}</a>
           </p>
           <div className="text-xs text-gray-600 flex gap-4 justify-center font-medium">
              <a href="https://www.englishclub.my.id" target="_blank" rel="noopener noreferrer" className="hover:text-gray-400 transition-colors">englishclub.my.id</a>
              <span>•</span>
              <a href="https://www.instagram.com/englishclubgresik/" target="_blank" rel="noopener noreferrer" className="hover:text-gray-400 transition-colors">@englishclubgresik</a>
           </div>
        </div>
      </div>
    </div>
  );
}

const KPICard = ({ title, value, subtext, icon: Icon, colorClass, onClick }: any) => (
    <Card onClick={onClick} className="cursor-pointer hover:-translate-y-1 transition-transform border-t-4 border-t-[#00D4FF] h-full">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0 flex-1">
          <p className="text-gray-400 text-xs sm:text-sm font-medium mb-1 truncate" title={title}>{title}</p>
          <h3 className="text-xl sm:text-2xl lg:text-3xl font-bold text-white mb-0.5 break-words leading-tight">{value}</h3>
          {subtext && <p className="text-xs text-gray-500 truncate" title={subtext}>{subtext}</p>}
        </div>
        <div className={`p-2.5 sm:p-3 rounded-xl shrink-0 ${colorClass}`}>
          <Icon size={22} className="text-white sm:w-6 sm:h-6" />
        </div>
      </div>
    </Card>
  );

// REPLACE THIS SECTION: DateTime Display Component
// Alasan: Mengurangi margin atas (mt-3) agar lebih merapat dengan teks sapaan dan ukuran tombol lebih ringkas di layar kecil.
const DateTimeDisplay = ({ language = 'en' }: any) => {
  const [now, setNow] = useState(new Date());

  useEffect(() => {
    // PERBAIKAN: Menghindari Memory Leak karena interval ganda dengan dependency array kosong []
    let interval;
    const secUntilNextMin = 60 - new Date().getSeconds();
    const timeout = setTimeout(() => {
      setNow(new Date());
      interval = setInterval(() => setNow(new Date()), 60000);
    }, secUntilNextMin * 1000);
    
    return () => {
      clearTimeout(timeout);
      if (interval) clearInterval(interval);
    };
  }, []);

  const dateOptions: Intl.DateTimeFormatOptions = { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' };
  const timeOptions: Intl.DateTimeFormatOptions = { hour: 'numeric', minute: '2-digit', hour12: true };

  return (
    <div className="flex flex-col sm:flex-row items-start sm:items-center gap-2.5 mt-3 text-gray-300 font-medium w-full">
      <div className="flex items-center gap-2 bg-[#0A0E17]/80 backdrop-blur-md px-3.5 py-2 rounded-xl border border-gray-700/50 text-xs sm:text-sm shadow-md hover:border-gray-600 transition-all w-full sm:w-auto">
        <CalendarIcon size={14} className="text-[#3B82F6]" />
        <span className="tracking-wide">{now.toLocaleDateString(language === 'id' ? 'id-ID' : 'en-US', dateOptions)}</span>
      </div>
      <div className="flex items-center gap-2 bg-[#00D4FF]/10 backdrop-blur-md text-[#00D4FF] px-3.5 py-2 rounded-xl border border-[#00D4FF]/20 text-xs sm:text-sm shadow-[0_4px_12px_rgba(0,212,255,0.05)] font-bold tracking-wider hover:bg-[#00D4FF]/15 transition-all w-full sm:w-auto">
        <Clock size={14} />
        <span>{now.toLocaleTimeString(language === 'id' ? 'id-ID' : 'en-US', timeOptions)}</span>
      </div>
    </div>
  );
};

// REPLACE THIS SECTION: Fetch data asli dari Open-Meteo & Layout Baru
// Alasan: Menghapus fixed min-width yang memaksa ruang kosong berlebih. Menggunakan w-full, h-full, dan justify-between agar menyesuaikan tinggi kontainer secara otomatis dan vertikalnya selaras.
const WeatherWidget = ({ language = 'en' }: any) => {
  // Nilai Default adalah cuaca rata-rata Gresik (agar tidak ada state kosong)
  const [weather, setWeather] = useState({ temp: 33, feelsLike: 36, humidity: 70, windSpeed: 15, code: 1, loading: true });

  useEffect(() => {
    let isMounted = true;
    
    const fetchWeather = async () => {
      try {
        // Pasang batas waktu (Timeout) 3 detik. Jika API lambat/diblokir, batalkan.
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 3000);
        
        const res = await fetch('https://api.open-meteo.com/v1/forecast?latitude=-7.169&longitude=112.641&current=temperature_2m,relative_humidity_2m,apparent_temperature,weather_code,wind_speed_10m', {
           signal: controller.signal
        });
        clearTimeout(timeoutId);
        
        if (!res.ok) throw new Error('API Error');
        const data = await res.json();
        
        if (isMounted) {
          setWeather({ 
            temp: Math.round(data.current.temperature_2m), 
            feelsLike: Math.round(data.current.apparent_temperature),
            humidity: Math.round(data.current.relative_humidity_2m),
            windSpeed: Math.round(data.current.wind_speed_10m),
            code: data.current.weather_code, 
            loading: false
          });
        }
      } catch (e) {
        // FALLBACK AMAN: Hentikan animasi loading dan tampilkan data default Gresik
        if (isMounted) {
          setWeather(prev => ({ ...prev, loading: false })); 
        }
      }
    };
    
    fetchWeather();
    const intervalId = setInterval(fetchWeather, 15 * 60 * 1000); 
    
    return () => {
      isMounted = false;
      clearInterval(intervalId);
    };
  }, []);

  if (weather.loading) {
    return (
       <div className="animate-pulse bg-gradient-to-br from-[#0A0E17]/90 to-[#151B26]/90 border border-gray-700/50 h-[180px] w-full xl:w-[320px] rounded-[24px] shadow-lg flex flex-col items-center justify-center gap-3">
          <Cloud className="text-gray-600 animate-bounce" size={32} />
          <span className="text-xs text-gray-600 font-bold tracking-widest">{language === 'id' ? 'MEMUAT...' : 'LOADING...'}</span>
       </div>
    );
  }

  // WMO Weather code mapping
  const getWeatherDetails = (code) => {
    if (code === 0) return { icon: Sun, label: language === 'id' ? 'Cerah' : 'Sunny', color: 'text-amber-400', bg: 'bg-amber-400/10' };
    if (code >= 1 && code <= 3) return { icon: Cloud, label: language === 'id' ? 'Berawan' : 'Cloudy', color: 'text-blue-300', bg: 'bg-blue-400/10' };
    if (code >= 51 && code <= 65) return { icon: CloudRain, label: language === 'id' ? 'Hujan' : 'Rainy', color: 'text-blue-500', bg: 'bg-blue-500/10' };
    if (code >= 95) return { icon: CloudLightning, label: language === 'id' ? 'Badai' : 'Stormy', color: 'text-purple-400', bg: 'bg-purple-500/10' };
    return { icon: Cloud, label: language === 'id' ? 'Berawan' : 'Cloudy', color: 'text-gray-300', bg: 'bg-gray-500/10' };
  };

  const { icon: Icon, label, color, bg } = getWeatherDetails(weather.code);

  return (
    <div className="flex flex-col justify-between bg-gradient-to-br from-[#0A0E17]/90 to-[#151B26]/90 backdrop-blur-xl border border-gray-700/50 rounded-[24px] shadow-[0_8px_32px_rgba(0,0,0,0.2)] p-4 sm:p-5 hover:border-[#00D4FF]/30 transition-all duration-300 w-full cursor-default relative overflow-hidden group h-full">
      {/* Decorative background glow */}
      <div className="absolute -top-10 -right-10 w-32 h-32 bg-[#00D4FF]/5 rounded-full blur-2xl group-hover:bg-[#00D4FF]/10 transition-colors pointer-events-none"></div>

      {/* Top Row: Main Temp & Location */}
      <div className="flex items-center justify-between mb-3 pb-3 border-b border-gray-700/50 relative z-10">
         <div className="flex items-center gap-3 sm:gap-4">
            <div className={`p-2.5 sm:p-3 rounded-2xl ${bg} border border-white/5 shadow-inner`}>
              <Icon size={24} className={color} />
            </div>
            <div className="flex flex-col">
              <div className="flex items-baseline gap-1">
                <span className="text-2xl sm:text-3xl font-black text-white leading-none tracking-tighter">{weather.temp}°<span className="text-lg sm:text-xl text-gray-400 font-bold">C</span></span>
              </div>
              <span className="text-xs sm:text-sm font-semibold text-[#00D4FF] uppercase tracking-widest mt-1">{label}</span>
            </div>
         </div>
         <div className="text-right pl-2">
            <span className="block text-[11px] sm:text-xs text-gray-400 font-semibold leading-tight">Kebomas,</span>
            <span className="block text-xs sm:text-sm text-gray-200 font-bold tracking-wide">Gresik</span>
         </div>
      </div>
      
      {/* Bottom Row: Detail Cards */}
      <div className="grid grid-cols-3 gap-2 relative z-10">
         <div className="flex flex-col items-center justify-center p-2 sm:p-2.5 bg-white/[0.03] hover:bg-white/[0.06] transition-colors rounded-xl border border-white/5">
            <Thermometer size={14} className="text-amber-400 mb-1" />
            <span className="text-[11px] text-gray-500 uppercase tracking-widest mb-0.5 font-bold">{language === 'id' ? 'Terasa' : 'Feels'}</span>
            <span className="text-[11px] sm:text-xs font-black text-white">{weather.feelsLike}°C</span>
         </div>
         <div className="flex flex-col items-center justify-center p-2 sm:p-2.5 bg-white/[0.03] hover:bg-white/[0.06] transition-colors rounded-xl border border-white/5">
            <Droplets size={14} className="text-blue-400 mb-1" />
            <span className="text-[11px] text-gray-500 uppercase tracking-widest mb-0.5 font-bold">{language === 'id' ? 'Lembap' : 'Humid'}</span>
            <span className="text-[11px] sm:text-xs font-black text-white">{weather.humidity}%</span>
         </div>
         <div className="flex flex-col items-center justify-center p-2 sm:p-2.5 bg-white/[0.03] hover:bg-white/[0.06] transition-colors rounded-xl border border-white/5">
            <Wind size={14} className="text-emerald-400 mb-1" />
            <span className="text-[11px] text-gray-500 uppercase tracking-widest mb-0.5 font-bold">{language === 'id' ? 'Angin' : 'Wind'}</span>
            <span className="text-[11px] sm:text-xs font-black text-white flex items-baseline gap-0.5">{weather.windSpeed}<span className="text-[11px] font-semibold text-gray-400">km/h</span></span>
         </div>
      </div>
    </div>
  );
};

const GreetingCard = ({ userName, children, isCloudConnected, language = 'en' }: any) => {
  const hour = new Date().getHours();
  let greeting = '';
  let emoji = '';
  let subtitle = '';

  if (hour >= 5 && hour < 11) {
    greeting = language === 'id' ? 'Selamat Pagi' : 'Good Morning';
    emoji = '☀️';
    subtitle = language === 'id' ? 'Awali harimu dengan semangat dan raih targetmu.' : 'Start your day with enthusiasm and achieve your goals.';
  } else if (hour >= 11 && hour < 15) {
    greeting = language === 'id' ? 'Selamat Siang' : 'Good Afternoon';
    emoji = '☁️';
    subtitle = language === 'id' ? 'Tetap semangat dan jaga produktivitasmu.' : 'Keep up the good work and stay productive.';
  } else if (hour >= 15 && hour < 18) {
    greeting = language === 'id' ? 'Selamat Sore' : 'Good Evening';
    emoji = '🌤️';
    subtitle = language === 'id' ? 'Selesaikan tugas hari ini dan bersiap untuk besok.' : "Finish today's tasks and prepare for tomorrow.";
  } else {
    greeting = language === 'id' ? 'Selamat Malam' : 'Good Night';
    emoji = '🌙';
    subtitle = language === 'id' ? 'Waktunya istirahat dan tinjau progres belajarmu.' : 'Take time to review your progress and recharge.';
  }

  const displayName = userName ? userName.split(' ')[0] : 'User';

  // REPLACE THIS SECTION: Perbaikan Rasio Responsif Desktop (~45%, ~35%, ~20%) dan pembungkus vertikal
  // Alasan: Membagi flex layout untuk Mobile (100%), Tablet (Baris 1: 55% + 40%, Baris 2: 100%), dan Desktop (45% + 35% + 20% satu baris). Memakai items-stretch agar tinggi selaras.
  return (
    <div className="w-full bg-[#151B26]/80 backdrop-blur-xl border border-[#00D4FF]/20 p-4 lg:p-5 rounded-[24px] shadow-[0_8px_32px_rgba(0,0,0,0.3)] animation-fade-in relative overflow-hidden flex flex-col md:flex-row md:flex-wrap lg:flex-nowrap justify-between items-stretch gap-3 lg:gap-4 mb-4 sm:mb-6 transition-all">
      <div className="absolute top-0 right-0 w-full h-full bg-gradient-to-bl from-[#00D4FF]/5 via-transparent to-transparent pointer-events-none"></div>
      <div className="absolute -bottom-20 -left-20 w-64 h-64 bg-[#3B82F6]/10 blur-[100px] rounded-full pointer-events-none"></div>
      
      {/* KIRI: Sapaan (~45% di Desktop, ~55% di Tablet) */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4 relative z-10 w-full md:w-[55%] lg:w-[45%]">
        <div className="flex-shrink-0 bg-[#0A0E17]/60 p-3.5 rounded-2xl border border-white/5 shadow-inner flex items-center justify-center">
          <span className="text-4xl drop-shadow-[0_0_20px_rgba(255,255,255,0.2)] select-none leading-none">
            {emoji}
          </span>
        </div>
        <div className="flex-1 w-full flex flex-col justify-center">
          <h2 className="text-2xl sm:text-3xl font-black text-white mb-1 tracking-tight leading-tight">
            {greeting}, <span className="bg-clip-text text-transparent bg-gradient-to-r from-[#00D4FF] to-[#3B82F6]">{displayName}</span>
          </h2>
          <p className="text-gray-400 font-medium text-xs sm:text-sm leading-relaxed max-w-sm lg:max-w-md">
            {subtitle}
          </p>
          <DateTimeDisplay language={language} />
        </div>
      </div>

      {/* TENGAH: Widget Cuaca (~35% di Desktop, ~40% di Tablet) */}
      <div className="relative z-10 w-full md:w-[40%] lg:w-[35%] shrink-0 flex">
         <WeatherWidget language={language} />
      </div>

      {/* KANAN: Cloud Status & Children (~20% di Desktop) */}
      <div className="relative z-10 w-full lg:w-[18%] lg:flex-1 flex flex-col md:flex-row lg:flex-col justify-center items-stretch md:items-center lg:items-end gap-3 shrink-0">
        
        {/* Status Cloud Real-time Badge */}
        <div className={`flex lg:flex-col xl:flex-row items-center gap-3 px-3.5 py-3 rounded-[20px] border backdrop-blur-md transition-all duration-500 shadow-md w-full justify-start xl:justify-center h-full sm:h-auto lg:h-full ${
            isCloudConnected 
            ? 'bg-emerald-500/10 border-emerald-500/20 shadow-[0_8px_20px_rgba(16,185,129,0.1)]' 
            : 'bg-rose-500/10 border-rose-500/20 shadow-[0_8px_20px_rgba(244,63,94,0.1)]'
        }`}>
           <div className="relative flex items-center justify-center flex-shrink-0">
              <div className={`p-2 rounded-xl ${isCloudConnected ? 'bg-emerald-500/20 text-emerald-400' : 'bg-rose-500/20 text-rose-400'}`}>
                 {isCloudConnected ? <Cloud size={20} /> : <CloudOff size={20} />}
              </div>
              <span className="absolute -bottom-0.5 -right-0.5 flex h-3 w-3">
                  {isCloudConnected && <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>}
                  <span className={`relative inline-flex rounded-full h-3 w-3 border-[2px] border-[#151B26] ${isCloudConnected ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
              </span>
           </div>
           <div className="flex flex-col lg:items-center xl:items-start text-left lg:text-center xl:text-left">
              <span className="text-[11px] text-gray-400 font-bold uppercase tracking-widest leading-none mb-1">{language === 'id' ? 'Status Sistem' : 'System Status'}</span>
              <span className={`text-xs sm:text-sm font-black leading-none flex items-center justify-start lg:justify-center xl:justify-start gap-1 tracking-wide ${isCloudConnected ? 'text-emerald-400' : 'text-rose-400'}`}>
                 {isCloudConnected ? (language === 'id' ? 'TERHUBUNG' : 'CONNECTED') : (language === 'id' ? 'MODE LURING' : 'OFFLINE MODE')}
              </span>
           </div>
        </div>

        {/* Children Render */}
        {children && (
          <div className="w-full">
            {children}
          </div>
        )}
      </div>
    </div>
  );
};

// KOMPONEN BARU: Radar Chart (Spider Web) Generator Tanpa Library Eksternal
// FIX #15: Generalize label positions agar tidak crash jika data.length !== 4
const RadarChart = ({ data, theme = 'dark' }: any) => {
  // Guard: butuh minimal 3 titik untuk membentuk polygon bermakna
  if (!Array.isArray(data) || data.length < 3) return null;

  const centerX = 50;
  const centerY = 50;
  const maxRadius = 35; // Ruang ekstra untuk teks label
  const n = data.length;

  // Distribusi sudut generik: titik pertama selalu di atas (- π/2)
  const getCoordinates = (value, index) => {
    const angle = (2 * Math.PI / n) * index - Math.PI / 2;
    const radius = ((value || 0) / 100) * maxRadius;
    return {
      x: centerX + radius * Math.cos(angle),
      y: centerY + radius * Math.sin(angle)
    };
  };

  // Titik label sedikit lebih jauh dari polygon agar tidak overlap
  const getLabelCoords = (index) => {
    const angle = (2 * Math.PI / n) * index - Math.PI / 2;
    const r = maxRadius + 8;
    return { x: centerX + r * Math.cos(angle), y: centerY + r * Math.sin(angle) };
  };

  const getLabelAnchor = (index) => {
    const angle = (2 * Math.PI / n) * index - Math.PI / 2;
    const cosA = Math.cos(angle);
    if (cosA > 0.3) return 'start';
    if (cosA < -0.3) return 'end';
    return 'middle';
  };

  const points = data.map((d: any, i: number) => getCoordinates(d.value || 0, i));
  const polygonPoints = points.map((p: any) => `${p.x},${p.y}`).join(' ');

  // Grid polygon generik (n sisi)
  const gridPolygonPoints = (r: number) =>
    Array.from({ length: n }, (_, i) => {
      const angle = (2 * Math.PI / n) * i - Math.PI / 2;
      return `${centerX + r * Math.cos(angle)},${centerY + r * Math.sin(angle)}`;
    }).join(' ');

  const strokeColor = theme === 'dark' ? '#00D4FF' : '#3B82F6';
  const fillColor = theme === 'dark' ? 'rgba(0, 212, 255, 0.2)' : 'rgba(59, 130, 246, 0.15)';
  const gridColor = theme === 'dark' ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';
  const textColor = theme === 'dark' ? '#9CA3AF' : '#64748B';
  const valueColor = theme === 'dark' ? '#FFFFFF' : '#0F172A';

  return (
    <svg viewBox="0 0 100 100" className="w-full h-full overflow-visible drop-shadow-sm">
      {/* Jaring Belakang */}
      {[20, 40, 60, 80, 100].map((val) => {
        const r = (val / 100) * maxRadius;
        return (
          <polygon
            key={val}
            points={gridPolygonPoints(r)}
            fill="none"
            stroke={gridColor}
            strokeWidth="0.5"
            strokeDasharray={val === 100 ? "0" : "1 1"}
          />
        );
      })}
      {/* Garis Sumbu dari Pusat ke Tiap Sudut */}
      {Array.from({ length: n }, (_, i) => {
        const angle = (2 * Math.PI / n) * i - Math.PI / 2;
        return (
          <line
            key={i}
            x1={centerX} y1={centerY}
            x2={centerX + maxRadius * Math.cos(angle)}
            y2={centerY + maxRadius * Math.sin(angle)}
            stroke={gridColor} strokeWidth="0.5"
          />
        );
      })}

      {/* Polygon Data Utama */}
      <polygon points={polygonPoints} fill={fillColor} stroke={strokeColor} strokeWidth="1.5" className="transition-all duration-1000 ease-out" />
      
      {/* Titik Sudut */}
      {points.map((p: any, i: number) => <circle key={i} cx={p.x} cy={p.y} r="1.5" fill={strokeColor} />)}

      {/* Teks Label Generik (bekerja untuk 3, 4, atau lebih titik) */}
      {data.map((d: any, i: number) => {
        const lc = getLabelCoords(i);
        const anchor = getLabelAnchor(i);
        return (
          <g key={i}>
            <text x={lc.x} y={lc.y - 1.5} textAnchor={anchor} fontSize="4.5" fill={textColor} fontWeight="bold">{d.label}</text>
            <text x={lc.x} y={lc.y + 4} textAnchor={anchor} fontSize="4" fill={valueColor} fontWeight="900">{d.value}</text>
          </g>
        );
      })}
    </svg>
  );
};

const StudentDashboard = ({ db, user, setActiveTab, today, isCloudConnected, language = 'en' }: any) => {
  const [currentTime, setCurrentTime] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const studentRecord = db.students.find(s => s.id === user.studentId) || { class: '-', level: '-', name: user.name };
  // FIX: pakai getStudentSession() agar sessionOverride dihormati, bukan hanya class
  const mySessionGroup = getStudentSession(studentRecord);
  const currentMonthPrefix = `${currentTime.getFullYear()}-${String(currentTime.getMonth() + 1).padStart(2, '0')}`;
  const currentMonthStr = String(currentTime.getMonth() + 1);
  const currentYearStr = String(currentTime.getFullYear());

  const myAttAll = db.studentAttendance.filter(a => a.studentId === user.studentId);
  const myAttThisMonth = myAttAll.filter(a => a.date.startsWith(currentMonthPrefix));
  const totalSessionsThisMonth = myAttThisMonth.length;
  const presentThisMonth = myAttThisMonth.filter(a => a.status === 'Present').length;
  const attRateThisMonth = totalSessionsThisMonth > 0 ? Math.round((presentThisMonth / totalSessionsThisMonth) * 100) : 0;

  const myAssessments = db.assessments.filter(a => a.studentId === user.studentId).sort((a, b) => {
     if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
     return Number(b.month) - Number(a.month);
  });
  const completedAssThisMonth = myAssessments.filter(a => Number(a.month) === Number(currentMonthStr) && String(a.year) === currentYearStr).length;
  const pendingAss = Math.max(0, 1 - completedAssThisMonth);

  // FIX #2: Gunakan fuzzy sessionMatches (konsisten dengan StudentReadOnlyJournalsModule)
  // agar angka jurnal di dashboard tidak berbeda dari daftar di halaman My Journals.
  const _sessionMatchesDash = (jGroup, sGroup) => {
    if (!jGroup || !sGroup) return false;
    if (jGroup === sGroup) return true;
    const a = jGroup.toLowerCase(), b = sGroup.toLowerCase();
    return a.includes(b) || b.includes(a);
  };
  const journalEntriesThisMonth = db.journals.filter(j => _sessionMatchesDash(j.sessionGroup, mySessionGroup) && j.date.startsWith(currentMonthPrefix)).length;

  const studentPlan = studentRecord.paymentPlan || 'Monthly';
  let paymentTarget = 0;
  if (studentPlan === 'Monthly') {
      // BUG FIX #5: Use fuzzy session match (consistent with other session comparisons in StudentDashboard)
      // to avoid paymentTarget being 0 when session strings have minor format differences.
      const _sessionMatchPayment = (cGroup, sGroup) => {
        if (!cGroup || !sGroup) return false;
        if (cGroup === sGroup) return true;
        const a = cGroup.toLowerCase(), b = sGroup.toLowerCase();
        return a.includes(b) || b.includes(a);
      };
      paymentTarget = db.calendar.filter(c => c.date.startsWith(currentMonthPrefix) && _sessionMatchPayment(c.sessionGroup || c.name, mySessionGroup)).length * 25000;
  } else {
      paymentTarget = db.studentAttendance.filter(a => a.studentId === user.studentId && a.date.startsWith(currentMonthPrefix) && a.status === 'Present').length * 25000;
  }
  
  const myPaymentsThisMonth = db.payments.filter(p => p.studentId === user.studentId && Number(p.month) === Number(currentMonthStr) && String(p.year) === currentYearStr && p.status === 'Paid');
  const totalPaidAmount = myPaymentsThisMonth.reduce((sum, p) => sum + Number(p.amount), 0);
  
  let paymentStatusText = '';
  if (paymentTarget === 0 && totalPaidAmount === 0) paymentStatusText = language === 'id' ? 'Belum Ada Tagihan' : 'No Target';
  else if (totalPaidAmount >= paymentTarget) paymentStatusText = language === 'id' ? 'Lunas' : 'Paid';
  else if (totalPaidAmount > 0) paymentStatusText = language === 'id' ? 'Parsial / Kurang' : 'Partial';
  else paymentStatusText = language === 'id' ? 'Belum Bayar' : 'Unpaid';
  
  const isPaid = totalPaidAmount >= paymentTarget && paymentTarget > 0;
  const isPartial = totalPaidAmount > 0 && totalPaidAmount < paymentTarget;

  const latestAss = myAssessments.length > 0 ? myAssessments[0] : null;
  const avgScore = latestAss ? latestAss.average : 0;
  const currentGrade = latestAss ? latestAss.grade : '-';

  // BUGFIX: Gunakan fuzzy match agar jadwal tetap muncul meski ada perbedaan format nama sesi kecil
  const upcomingClasses = db.calendar
      .filter(c => {
        if ((c.date || '') < today) return false;
        const cGroup = (c.sessionGroup || c.name || '').toLowerCase();
        const sGroup = (mySessionGroup || '').toLowerCase();
        return cGroup === sGroup || cGroup.includes(sGroup) || sGroup.includes(cGroup);
      })
      .sort((a,b) => String(a.date || '').localeCompare(String(b.date || '')))
      .slice(0, 5);
  const nextClass = upcomingClasses.length > 0 ? upcomingClasses[0] : null;

  const scores = latestAss && latestAss.scores ? latestAss.scores : {};
  const getScore = (subject) => Number(scores[subject]) || 0;
  
  const isKindergarten = studentRecord.level === 'Kindergarten' || ['PAUD', 'TK A', 'TK B'].includes(studentRecord.class);

  const skillProgress = isKindergarten ? [
     { label: 'Reading', value: getScore('Reading'), color: 'bg-blue-500' },
     { label: 'Writing', value: getScore('Writing'), color: 'bg-purple-500' },
     { label: 'Math', value: getScore('Math'), color: 'bg-emerald-500' },
     { label: 'English', value: getScore('English'), color: 'bg-amber-500' }
  ] : [
     { label: 'Speaking', value: getScore('Speaking'), color: 'bg-blue-500' },
     { label: 'Writing', value: getScore('Writing'), color: 'bg-purple-500' },
     { label: 'Reading', value: getScore('Reading'), color: 'bg-emerald-500' },
     { label: 'Listening', value: getScore('Listening'), color: 'bg-amber-500' }
  ];

  return (
     <div className="space-y-5 sm:space-y-6 w-full max-w-full overflow-hidden animation-fade-in pb-6 font-sans">
        
        <GreetingCard userName={user?.name} isCloudConnected={isCloudConnected} language={language} />

        {/* TOP HERO SECTION */}
        <div className="bg-gradient-to-br from-[#0A3D91] to-[#051126] border border-[#00D4FF]/30 p-4 sm:p-7 rounded-[20px] shadow-2xl relative overflow-hidden flex flex-col md:flex-row justify-between items-start md:items-center gap-4 sm:gap-6 -mt-2">
           <div className="absolute top-[-20%] right-[-10%] w-[60%] h-[150%] bg-[#00D4FF]/10 blur-[120px] rounded-full pointer-events-none"></div>
           
           <div className="relative z-10 w-full md:w-2/3">
              <h2 className="text-xl sm:text-2xl font-bold text-white mb-1.5 tracking-tight">{language === 'id' ? 'Ringkasan Akademik Anda' : 'Your Academic Overview'}</h2>
              <p className="text-blue-100/80 font-medium text-sm mb-4">{language === 'id' ? 'Apa yang harus saya lakukan hari ini?' : 'What should I do today?'}</p>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                 <div>
                    <p className="text-xs text-blue-200/60 uppercase tracking-widest font-bold mb-1.5 flex items-center gap-1.5"><Clock size={12}/> {language === 'id' ? 'Kelas Berikutnya' : 'Next Class'}</p>
                    <p className="text-white font-semibold text-sm line-clamp-1">{nextClass ? `${nextClass.sessionGroup || nextClass.name} @ ${nextClass.startTime}` : (language === 'id' ? 'Tidak ada hari ini' : 'None today')}</p>
                 </div>
                 <div>
                    <p className="text-xs text-blue-200/60 uppercase tracking-widest font-bold mb-1.5 flex items-center gap-1.5"><CheckSquare size={12}/> {language === 'id' ? 'Tugas' : 'Assignments'}</p>
                    <p className="text-white font-semibold text-sm">{pendingAss} {language === 'id' ? 'Tertunda' : 'Pending'}</p>
                 </div>
                 <div>
                    <p className="text-xs text-blue-200/60 uppercase tracking-widest font-bold mb-1.5 flex items-center gap-1.5"><UserCheck size={12}/> {language === 'id' ? 'Kehadiran' : 'Attendance'}</p>
                    <p className="text-white font-semibold text-sm">{attRateThisMonth}%</p>
                 </div>
                 <div>
                    <p className="text-xs text-blue-200/60 uppercase tracking-widest font-bold mb-1.5 flex items-center gap-1.5"><Award size={12}/> {language === 'id' ? 'Nilai Saat Ini' : 'Current Score'}</p>
                    <p className="text-white font-semibold text-sm">{avgScore || '-'} <span className="text-blue-300 font-medium text-xs ml-1">({language === 'id' ? 'Peringkat' : 'Grade'} {currentGrade})</span></p>
                 </div>
              </div>
           </div>

           <div className="relative z-10 w-full md:w-auto flex flex-row md:flex-col gap-2 min-w-[180px]">
              <Button onClick={() => setActiveTab('calendar')} className="flex-1 md:flex-none w-full min-h-[42px] shadow-[0_0_20px_rgba(0,212,255,0.25)] text-sm font-bold rounded-xl hover:scale-105 transition-transform" icon={CalendarIcon}>{language === 'id' ? 'Gabung Kelas' : 'Join Class'}</Button>
              <Button onClick={() => setActiveTab('my_assessments')} variant="secondary" className="flex-1 md:flex-none w-full min-h-[42px] text-sm font-bold rounded-xl border-[#00D4FF]/30 hover:bg-[#00D4FF]/10 text-[#00D4FF]" icon={FileText}>{language === 'id' ? 'Lihat Tugas' : 'View Assignments'}</Button>
           </div>
        </div>

        {/* QUICK STATS CARDS */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 sm:gap-4">
           <KPICard title={language === 'id' ? 'Kehadiran' : 'Attendance'} value={`${attRateThisMonth}%`} subtext={`${presentThisMonth} ${language === 'id' ? 'sesi dihadiri' : 'sessions attended'}`} icon={UserCheck} colorClass={attRateThisMonth >= 80 ? "bg-emerald-500" : "bg-amber-500"} onClick={() => setActiveTab('my_attendance')} />
           <KPICard title={language === 'id' ? 'Tugas' : 'Assignments'} value={pendingAss} subtext={pendingAss > 0 ? (language === 'id' ? "Batas waktu bulan ini" : "Due this month") : (language === 'id' ? "Semua selesai!" : "All caught up!")} icon={CheckSquare} colorClass={pendingAss === 0 ? "bg-emerald-500" : "bg-rose-500"} onClick={() => setActiveTab('my_assessments')} />
           <KPICard title={language === 'id' ? 'Jurnal Belajar' : 'Learning Journal'} value={journalEntriesThisMonth} subtext={language === 'id' ? 'Total entri' : "Total entries"} icon={BookOpen} colorClass="bg-purple-500" onClick={() => setActiveTab('my_journals')} />
           <KPICard title={language === 'id' ? 'Nilai Akademik' : 'Academic Score'} value={avgScore || '-'} subtext={`${language === 'id' ? 'Peringkat' : 'Grade'}: ${currentGrade}`} icon={Award} colorClass="bg-indigo-500" onClick={() => setActiveTab('my_report')} />
           <KPICard title={language === 'id' ? 'Status SPP' : 'Payment Status'} value={paymentStatusText} subtext={paymentTarget > 0 ? `Target: Rp ${paymentTarget.toLocaleString()}` : (language === 'id' ? 'Biaya SPP' : "Tuition fee")} icon={DollarSign} colorClass={isPaid ? "bg-emerald-500" : isPartial ? "bg-yellow-500" : "bg-rose-500"} onClick={() => setActiveTab('my_payments')} />
        </div>

        {/* TODAY'S ACTIVITY */}
        <div>
           <h3 className="text-xl font-bold text-white mb-4 flex items-center gap-2"><Activity size={22} className="text-[#00D4FF]" /> {language === 'id' ? 'Activities Hari Ini' : "Today's Activity"}</h3>
           <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
              {/* Next Class */}
              <Card className="bg-[#151B26]/80 backdrop-blur-md border-t-4 border-t-emerald-400 p-4 sm:p-5 hover:-translate-y-1 transition-transform shadow-lg flex flex-col">
                 <div className="flex justify-between items-start mb-4">
                    <div className="p-2.5 bg-emerald-500/10 rounded-xl"><CalendarIcon size={22} className="text-emerald-400"/></div>
                    <span className="px-2.5 py-0.5 text-[11px] font-semibold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-full uppercase tracking-widest">{language === 'id' ? 'Selanjutnya' : 'Up Next'}</span>
                 </div>
                 <h4 className="font-bold text-white text-base mb-1">{nextClass ? (nextClass.sessionGroup || nextClass.name) : (language === 'id' ? 'Tidak ada kelas hari ini' : 'No classes today')}</h4>
                 <p className="text-sm text-gray-400 mb-3 flex-grow">{nextClass ? `${nextClass.startTime} - ${nextClass.endTime} • ${nextClass.tutor}` : (language === 'id' ? 'Istirahat dan tinjau catatan Anda.' : 'Take a break and review your notes.')}</p>
                 <Button variant="secondary" className="w-full text-sm font-semibold py-2 bg-emerald-500/10 border-none text-emerald-400 hover:bg-emerald-500/20 rounded-lg" disabled={!nextClass} onClick={() => setActiveTab('calendar')}>{language === 'id' ? 'Lihat Jadwal' : 'View Schedule'}</Button>
              </Card>

              {/* Assignments Due */}
              <Card className="bg-[#151B26]/80 backdrop-blur-md border-t-4 border-t-rose-400 p-4 sm:p-5 hover:-translate-y-1 transition-transform shadow-lg flex flex-col">
                 <div className="flex justify-between items-start mb-4">
                    <div className="p-2.5 bg-rose-500/10 rounded-xl"><CheckSquare size={22} className="text-rose-400"/></div>
                    {pendingAss > 0 && <span className="px-2.5 py-0.5 text-[11px] font-semibold bg-rose-500/10 text-rose-400 border border-rose-500/20 rounded-full uppercase tracking-widest">{language === 'id' ? 'Perlu Tindakan' : 'Action Needed'}</span>}
                 </div>
                 <h4 className="font-bold text-white text-base mb-1">{pendingAss > 0 ? (language === 'id' ? `${pendingAss} Tugas Tertunda` : `${pendingAss} Pending Assessment`) : (language === 'id' ? 'Semua Selesai!' : 'All Caught Up!')}</h4>
                 <p className="text-sm text-gray-400 mb-3 flex-grow">{pendingAss > 0 ? (language === 'id' ? 'Ada tugas yang harus diselesaikan bulan ini.' : 'You have assessments due this month.') : (language === 'id' ? 'Tidak ada tugas saat ini.' : 'No pending assignments at the moment.')}</p>
                 <Button variant="secondary" className="w-full text-sm font-semibold py-2 bg-rose-500/10 border-none text-rose-400 hover:bg-rose-500/20 rounded-lg" onClick={() => setActiveTab('my_assessments')}>{language === 'id' ? 'Buka Tugas' : 'Go to Assessments'}</Button>
              </Card>

              {/* Journal Reminder */}
              <Card className="bg-[#151B26]/80 backdrop-blur-md border-t-4 border-t-purple-400 p-4 sm:p-5 hover:-translate-y-1 transition-transform shadow-lg flex flex-col">
                 <div className="flex justify-between items-start mb-4">
                    <div className="p-2.5 bg-purple-500/10 rounded-xl"><BookOpen size={22} className="text-purple-400"/></div>
                 </div>
                 <h4 className="font-bold text-white text-base mb-1">{language === 'id' ? 'Jurnal Belajar' : 'Learning Journal'}</h4>
                 <p className="text-sm text-gray-400 mb-3 flex-grow">{language === 'id' ? 'Tinjau materi terbaru dari tutor Anda.' : 'Review the latest materials from your tutor.'}</p>
                 <Button variant="secondary" className="w-full text-sm font-semibold py-2 bg-purple-500/10 border-none text-purple-400 hover:bg-purple-500/20 rounded-lg" onClick={() => setActiveTab('my_journals')}>{language === 'id' ? 'Buka Jurnal' : 'Open Journal'}</Button>
              </Card>
           </div>
        </div>

        {/* UPCOMING CLASSES */}
        <div>
           <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold text-white flex items-center gap-2"><CalendarIcon size={22} className="text-emerald-400" /> {language === 'id' ? 'Kelas Mendatang' : 'Upcoming Classes'}</h3>
              <Button variant="ghost" className="text-xs px-4 py-2 bg-[#151B26] hover:bg-gray-800 rounded-lg" onClick={() => setActiveTab('calendar')}>{language === 'id' ? 'Lihat Semua' : 'See All'}</Button>
           </div>
           <div className="flex overflow-x-auto gap-4 pb-4 custom-scrollbar snap-x">
              {upcomingClasses.length > 0 ? upcomingClasses.map(c => (
                 <Card key={c.id} className="min-w-[280px] sm:min-w-[320px] bg-[#151B26] p-5 border border-gray-800 shadow-md hover:border-emerald-500/40 hover:shadow-[0_8px_30px_rgba(16,185,129,0.1)] transition-all snap-start">
                    <div className="flex justify-between items-start mb-4">
                        <div className="flex flex-col gap-1.5">
                           <span className="px-3 py-1.5 text-xs font-bold bg-[#0B0F19] text-emerald-400 border border-emerald-500/20 rounded-lg tracking-wide">{safeDateDisplay(c.date, language === 'id' ? 'id-ID' : 'en-US', {weekday: 'short', month:'short', day:'numeric'})}</span>
                           <span className="text-[11px] font-semibold text-amber-400 uppercase tracking-widest">{calculateDaysLeft(c.date, today)}</span>
                        </div>
                        <span className="text-xs font-semibold text-gray-400 bg-gray-800/50 px-2.5 py-1.5 rounded-lg flex items-center gap-1.5"><Clock size={12}/> {c.startTime}</span>
                    </div>
                    <h4 className="font-bold text-white text-lg mb-2 truncate">{c.sessionGroup || c.name}</h4>
                    <div className="flex items-center gap-2 text-sm text-gray-400">
                        <User size={16} className="text-purple-400" />
                        <span className="truncate">{c.tutor}</span>
                    </div>
                 </Card>
              )) : (
                 <div className="w-full p-8 text-center bg-[#151B26] rounded-xl border border-gray-800">
                    <CalendarIcon size={40} className="mx-auto text-gray-700 mb-3" />
                    <p className="text-gray-400 font-medium text-sm">{language === 'id' ? 'Tidak ada jadwal kelas mendatang.' : 'No upcoming classes scheduled.'}</p>
                 </div>
              )}
           </div>
        </div>

        {/* LEARNING PROGRESS & ANNOUNCEMENTS */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
           <Card className="bg-[#151B26] border-gray-800 shadow-lg p-6 sm:p-8">
              <div className="flex items-center gap-3 mb-8">
                  <div className="p-2.5 bg-blue-500/10 rounded-lg border border-blue-500/20"><BarChart3 size={24} className="text-blue-400"/></div>
                  <h3 className="text-xl font-bold text-white">{language === 'id' ? 'Progres Belajar' : 'Learning Progress'}</h3>
              </div>
              <div className="w-full flex justify-center items-center min-h-[220px]">
                 {latestAss ? (
                    <div className="w-full max-w-[280px] aspect-square">
                       <RadarChart data={skillProgress} theme="dark" />
                    </div>
                 ) : (
                    <div className="flex flex-col items-center justify-center py-10">
                       <Award size={48} className="text-gray-700 mb-4" />
                       <p className="text-sm text-gray-500 font-medium">{language === 'id' ? 'Belum ada data nilai.' : 'No assessment data available yet.'}</p>
                    </div>
                 )}
              </div>
           </Card>

           <Card className="bg-[#151B26] border-gray-800 shadow-lg p-0 flex flex-col overflow-hidden">
              <div className="p-6 border-b border-gray-800 flex justify-between items-center bg-[#0A0E17]">
                 <h3 className="text-xl font-bold text-white flex items-center gap-3"><div className="p-2 bg-yellow-500/10 rounded-lg border border-yellow-500/20"><Bell size={20} className="text-yellow-400"/></div> {language === 'id' ? 'Pengumuman' : 'Announcements'}</h3>
                 <Button variant="ghost" className="text-xs py-2 px-4 bg-[#151B26] hover:bg-gray-800 rounded-lg" onClick={() => setActiveTab('announcements')}>{language === 'id' ? 'Lihat Semua' : 'View All'}</Button>
              </div>
              <div className="p-6 space-y-4 flex-1">
                 {db.announcements.slice(-3).reverse().map(a => (
                    <div key={a.id} className="p-5 bg-[#0B0F19] rounded-[16px] border border-gray-800 hover:border-gray-700 hover:shadow-lg transition-all group">
                       <h4 className="font-bold text-white group-hover:text-[#00D4FF] transition-colors text-base mb-1.5">{a.title}</h4>
                       <p className="text-[11px] text-gray-500 mb-3 font-semibold uppercase tracking-wider">{a.date} • {a.author}</p>
                       <p className="text-sm text-gray-400 line-clamp-2 leading-relaxed">{a.content}</p>
                    </div>
                 ))}
                 {db.announcements.length === 0 && <p className="text-sm text-gray-500 text-center py-10">{language === 'id' ? 'Tidak ada pengumuman.' : 'No announcements available.'}</p>}
              </div>
           </Card>
        </div>
     </div>
  );
};

const AdminDashboard = ({ db, setDb, user, setActiveTab, today, isCloudConnected, language = 'en', showToast }: any) => {
  const [leaderboardFilter, setLeaderboardFilter] = useState('All');
  const [expModalStudent, setExpModalStudent] = useState(null);
  const [expInput, setExpInput] = useState('');
  
  const getActiveCount = (collection) => (db[collection] || []).filter((item) => item.status === 'Active' || item.active === 'Active').length;
  
  const totalStudents = getActiveCount('students');
  const totalTutors = getActiveCount('tutors');
  const studentAttToday = db.studentAttendance.filter(a => a.date === today && a.status === 'Present').length;
  const tutorCheckInToday = db.tutorAttendance.filter(a => a.date === today && a.status === 'Present').length;
  
  const dObj = new Date();
  const currentMonth = String(dObj.getMonth() + 1);
  const currentYear = String(dObj.getFullYear());
  
  const currentJournals = db.journals.filter(j => j.date.startsWith(`${currentYear}-${currentMonth.padStart(2, '0')}`)).length;
  const paidPayrollCount = db.payroll.filter(p => Number(p.month) === Number(currentMonth) && String(p.year) === String(currentYear) && p.status === 'Paid').length;

  const activeStudentIds = db.students.filter(s => s.status === 'Active' || s.active === 'Active').map(s => s.id);
  const currentMonthPayments = db.payments.filter(p => Number(p.month) === Number(currentMonth) && String(p.year) === String(currentYear) && activeStudentIds.includes(p.studentId));
  const totalRevenueAmount = currentMonthPayments.filter(p => p.status === 'Paid').reduce((sum, p) => sum + Number(p.amount), 0);
  const paidInvoicesCount = currentMonthPayments.filter(p => p.status === 'Paid').length;

  // Calculate Expected Revenue globally accounting for Payment Plans
  const currentMonthPrefix = `${currentYear}-${currentMonth.padStart(2, '0')}`;
  let expectedRevenueAmount = 0;
  let fullyPaidStudentsCount = 0;
  const activeStudents = db.students.filter(s => s.status === 'Active' || s.active === 'Active');
  
  activeStudents.forEach(s => {
     // FREE plan: tidak masuk hitungan expected revenue sama sekali
     if (s.paymentPlan === 'Free') return;
     const sGroup = getStudentSession(s);
     // Semua siswa (Monthly maupun Per Visit) dihitung berdasarkan jumlah jadwal bulan ini
     const scheduledCount = db.calendar.filter(c => c.date.startsWith(currentMonthPrefix) && (c.sessionGroup || c.name) === sGroup).length;
     const studentTarget = scheduledCount * 25000;
     expectedRevenueAmount += studentTarget;

     const studentPaid = currentMonthPayments.filter(p => p.studentId === s.id && p.status === 'Paid').reduce((sum, p) => sum + Number(p.amount), 0);
     if (studentTarget > 0 && studentPaid >= studentTarget) fullyPaidStudentsCount++;
  });

  const totalOutstandingRevenueAmount = Math.max(0, expectedRevenueAmount - totalRevenueAmount);

  const sessionRevenueData = SESSIONS.map(session => {
      const sessionStudents = db.students.filter(s => s.status === 'Active' && getStudentSession(s) === session).length;
      const sessionPayments = currentMonthPayments.filter(p => p.sessionGroup === session && p.status === 'Paid');
      const sessionCollected = sessionPayments.reduce((sum, p) => sum + Number(p.amount), 0);
      const sessionPaidCount = sessionPayments.length;
      const sessionOutstanding = sessionStudents - sessionPaidCount;
      
      return { session, students: sessionStudents, collected: sessionPaidCount, outstanding: sessionOutstanding, revenue: sessionCollected };
  });

  const upcomingCalendar = db.calendar
      .filter(c => (c.date || '') >= today)
      .sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')))
      .slice(0, 5);

  return (
    <div className="space-y-4 sm:space-y-6 animation-fade-in w-full max-w-full overflow-hidden font-sans">
      <GreetingCard userName={user?.name} isCloudConnected={isCloudConnected} language={language} />

      {/* 8 Main Stats */}
      <div className="flex flex-wrap justify-center gap-3">
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Total Siswa' : 'Total Students'} value={totalStudents} icon={Users} colorClass="bg-blue-500" onClick={() => setActiveTab('students')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Total Tutor' : 'Total Tutors'} value={totalTutors} icon={Briefcase} colorClass="bg-purple-500" onClick={() => setActiveTab('tutors')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Hadir Siswa (Hari Ini)' : 'Student Att. Today'} value={studentAttToday} icon={UserCheck} colorClass="bg-emerald-500" onClick={() => setActiveTab('student_attendance')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Check-In Tutor' : 'Tutor Check-In'} value={tutorCheckInToday} icon={CheckSquare} colorClass="bg-teal-500" onClick={() => setActiveTab('tutor_attendance')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Jurnal (Bulan)' : 'Journals (Mo)'} value={currentJournals} icon={BookOpen} colorClass="bg-yellow-500" onClick={() => setActiveTab('journals')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Gaji Dibayar' : 'Paid Payroll'} value={paidPayrollCount} icon={FileText} colorClass="bg-rose-500" onClick={() => setActiveTab('payroll')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Menunggu Bayar' : 'Pending Payments'} value={Math.max(0, totalStudents - fullyPaidStudentsCount)} subtext={language === 'id' ? 'Siswa bulan ini' : 'Students this month'} icon={AlertCircle} colorClass={totalStudents - fullyPaidStudentsCount > 0 ? "bg-orange-500" : "bg-emerald-500"} onClick={() => setActiveTab('payments')} /></div>
         <div className="w-[calc(50%-0.375rem)] sm:w-[calc(33.333%-0.5rem)] xl:w-[calc(25%-0.5625rem)]"><KPICard title={language === 'id' ? 'Materi/Tugas' : 'Materials'} value={(db.materials || []).length} subtext={language === 'id' ? 'Total diposting' : 'Total posted'} icon={Database} colorClass="bg-cyan-500" onClick={() => setActiveTab('materials')} /></div>
      </div>

      {/* Total Revenue Cards */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 sm:gap-6">
         <Card className="bg-gradient-to-br from-amber-500/10 to-[#151B26] border-amber-500/20 hover:-translate-y-1 transition-transform cursor-pointer p-5 sm:p-6" onClick={() => setActiveTab('payments')}>
           <div className="flex items-center gap-4">
              <div className="p-3 sm:p-4 bg-amber-500/20 rounded-xl shrink-0"><DollarSign size={24} className="text-amber-400"/></div>
              <div className="min-w-0 flex-1">
                 <p className="text-xs sm:text-sm text-gray-400 font-medium truncate">{language === 'id' ? 'Pendapatan Tertunda' : 'Outstanding Revenue'}</p>
                 <h3 className="text-lg sm:text-2xl font-bold text-white break-words leading-tight">Rp {totalOutstandingRevenueAmount.toLocaleString('id-ID')}</h3>
                 <p className="text-xs text-amber-500/80 font-medium mt-0.5 truncate">{language === 'id' ? 'Target - Terkumpul' : 'Expected - Collected'}</p>
              </div>
           </div>
         </Card>
         <Card className="bg-gradient-to-br from-blue-500/10 to-[#151B26] border-blue-500/20 hover:-translate-y-1 transition-transform cursor-pointer p-5 sm:p-6" onClick={() => setActiveTab('payments')}>
           <div className="flex items-center gap-4">
              <div className="p-3 sm:p-4 bg-blue-500/20 rounded-xl shrink-0"><CheckCircle2 size={24} className="text-blue-400"/></div>
              <div className="min-w-0 flex-1">
                 <p className="text-xs sm:text-sm text-gray-400 font-medium truncate">{language === 'id' ? 'Siswa Lunas' : 'Fully Paid Students'}</p>
                 <h3 className="text-lg sm:text-2xl font-bold text-white break-words leading-tight">{fullyPaidStudentsCount} {language === 'id' ? 'Siswa' : 'Students'}</h3>
              </div>
           </div>
         </Card>
         <Card className="bg-gradient-to-br from-purple-500/10 to-[#151B26] border-purple-500/20 p-5 sm:p-6">
           <div className="flex items-center gap-4">
              <div className="p-3 sm:p-4 bg-purple-500/20 rounded-xl shrink-0"><Activity size={24} className="text-purple-400"/></div>
              <div className="min-w-0 flex-1">
                 <p className="text-xs sm:text-sm text-gray-400 font-medium truncate">{language === 'id' ? 'Progres Penagihan' : 'Collection Progress'}</p>
                 <h3 className="text-lg sm:text-2xl font-bold text-white break-words leading-tight">{expectedRevenueAmount > 0 ? Math.round((totalRevenueAmount / expectedRevenueAmount) * 100) : 0}{language === 'id' ? '% Terkumpul' : '% Collected'}</h3>
              </div>
           </div>
         </Card>
      </div>

      {/* Revenue by Session Table & Big Card */}
      <Card className="p-0 overflow-hidden border-[#00D4FF]/20 shadow-lg">
         <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
            <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><BarChart3 size={20} className="text-[#00D4FF]" /> {language === 'id' ? 'Pendapatan Sesi' : 'Revenue by Session'} <span className="text-xs text-gray-400 ml-1 sm:ml-2 font-medium">({MONTHS[Number(currentMonth)-1]} {currentYear})</span></h3>
         </div>
         <div className="w-full overflow-x-auto custom-scrollbar">
           <table className="w-full text-left text-sm min-w-[600px] whitespace-nowrap">
             <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-xs font-semibold">
               <tr><th className="p-3 sm:p-4 text-center w-12 text-gray-400">No.</th><th className="p-3 sm:p-4">{language === 'id' ? 'Sesi' : 'Session'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Siswa' : 'Students'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Terkumpul (Lunas)' : 'Collected (Paid)'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Tertunda' : 'Outstanding'}</th><th className="p-3 sm:p-4 text-right">{language === 'id' ? 'Pendapatan' : 'Revenue'}</th></tr>
             </thead>
             <tbody className="divide-y divide-gray-800">
               {sessionRevenueData.map((row, idx) => (
                  <tr key={row.session} className="hover:bg-[#0A0E17] transition-colors">
                    <td className="p-3 sm:p-4 text-center text-gray-500 font-medium">{idx + 1}</td>
                    <td className="p-3 sm:p-4 font-bold text-white text-xs sm:text-sm">{row.session}</td>
                    <td className="p-3 sm:p-4 text-center text-gray-300">{row.students}</td>
                    <td className="p-3 sm:p-4 text-center text-emerald-400 font-bold">{row.collected}</td>
                    <td className="p-3 sm:p-4 text-center text-rose-400 font-bold">{row.outstanding}</td>
                    <td className="p-3 sm:p-4 text-right font-black text-[#00D4FF] text-sm sm:text-base">Rp {row.revenue.toLocaleString()}</td>
                  </tr>
               ))}
             </tbody>
           </table>
         </div>
         <div className="bg-gradient-to-r from-[#0A3D91] to-[#051126] p-5 sm:p-8 flex flex-col sm:flex-row justify-between items-center sm:items-center border-t border-[#00D4FF]/30 gap-3 sm:gap-0">
            <div className="text-center sm:text-left w-full sm:w-auto">
               <p className="text-[#00D4FF] font-black tracking-widest uppercase text-xs sm:text-sm md:text-base drop-shadow-md">{language === 'id' ? 'TOTAL PENDAPATAN SEMUA SESI' : 'TOTAL REVENUE ALL SESSIONS'}</p>
               <p className="text-blue-200/60 text-xs mt-1 font-medium">{language === 'id' ? 'Akumulasi dari tagihan lunas bulan ini' : 'Accumulated from paid invoices this month'}</p>
            </div>
            <div className="text-2xl sm:text-4xl md:text-5xl font-black text-white tracking-tight drop-shadow-lg text-center sm:text-right w-full sm:w-auto break-words">
               Rp {totalRevenueAmount.toLocaleString()}
            </div>
         </div>
      </Card>

      {/* 2 Columns: Announcements & Calendar */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-6 w-full">
         <Card className="border-t-4 border-t-yellow-400 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><Bell size={18} className="text-yellow-400"/> {language === 'id' ? 'Pengumuman Terbaru' : 'Recent Announcements'}</h3>
               <Button variant="ghost" className="text-xs py-2 px-3 min-h-[36px]" onClick={() => setActiveTab('announcements')}>{language === 'id' ? 'Lihat Semua' : 'View All'}</Button>
            </div>
            <div className="p-4 sm:p-5 space-y-3 sm:space-y-4 bg-[#151B26]">
               {db.announcements.slice(-3).reverse().map(a => (
                  <div key={a.id} className="p-3 sm:p-4 bg-[#0B0F19] rounded-xl border border-gray-800 hover:border-gray-700 transition-colors">
                     <h4 className="font-bold text-[#00D4FF] text-sm sm:text-base mb-1">{a.title}</h4>
                     <p className="text-xs text-gray-500 mb-1.5 sm:mb-2 font-medium">{a.date}</p>
                     <p className="text-xs sm:text-sm text-gray-300 line-clamp-2">{a.content}</p>
                  </div>
               ))}
               {db.announcements.length === 0 && <p className="text-xs sm:text-sm text-gray-500 text-center py-4">{language === 'id' ? 'Tidak ada pengumuman.' : 'No announcements.'}</p>}
            </div>
         </Card>

         <Card className="border-t-4 border-t-emerald-400 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><CalendarIcon size={18} className="text-emerald-400"/> {language === 'id' ? 'Kalender Mendatang' : 'Upcoming Calendar'}</h3>
               <Button variant="ghost" className="text-xs py-2 px-3 min-h-[36px]" onClick={() => setActiveTab('calendar')}>{language === 'id' ? 'Kalender Lengkap' : 'Full Calendar'}</Button>
            </div>
            <div className="p-4 sm:p-5 space-y-3 bg-[#151B26]">
               {upcomingCalendar.map(c => (
                  <div key={c.id} className="flex justify-between items-center p-3 sm:p-4 bg-[#0B0F19] rounded-xl border border-gray-800 border-l-4 border-l-emerald-500 hover:bg-[#0A0E17] transition-colors gap-3">
                     <div className="flex-1 min-w-0">
                        <p className="font-bold text-white text-sm sm:text-base truncate">{c.sessionGroup || c.name}</p>
                        <p className="text-xs text-blue-400 font-medium truncate mt-0.5">{c.tutor}</p>
                        <div className="flex items-center flex-wrap gap-2 mt-1.5 text-xs">
                           <span className="text-emerald-400 font-medium whitespace-nowrap">{safeDateDisplay(c.date, language === 'id' ? 'id-ID' : 'en-US', {weekday: 'short'})}, {c.date}</span>
                           <span className="text-gray-500 hidden sm:inline whitespace-nowrap">{c.startTime} - {c.endTime}</span>
                           <span className="text-amber-400 font-bold bg-amber-400/10 px-2 py-0.5 rounded ml-auto">{calculateDaysLeft(c.date, today)}</span>
                        </div>
                     </div>
                  </div>
               ))}
               {upcomingCalendar.length === 0 && <p className="text-xs sm:text-sm text-gray-500 text-center py-4">{language === 'id' ? 'Tidak ada jadwal mendatang.' : 'No upcoming events scheduled.'}</p>}
            </div>
         </Card>
      </div>

      {/* NEW: GLOBAL LEADERBOARD SECTION (ADMIN ONLY) */}
      <Card className="p-0 overflow-hidden border-t-4 border-t-yellow-400 shadow-lg mt-4 sm:mt-6">
         <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
            <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><Trophy size={20} className="text-yellow-400" /> {language === 'id' ? 'Peringkat Siswa Global' : 'Global Student Leaderboard'}</h3>
            <select
              value={leaderboardFilter}
              onChange={(e) => setLeaderboardFilter(e.target.value)}
              className="bg-[#0B0F19] border border-gray-700 text-gray-300 rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-yellow-400 appearance-none cursor-pointer w-full sm:w-auto"
              style={{ backgroundImage: 'url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%2394A3B8%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E")', backgroundRepeat: 'no-repeat', backgroundPosition: 'right 0.5rem top 50%', backgroundSize: '0.65rem auto', paddingRight: '2rem' }}
            >
              <option value="All">{language === 'id' ? 'All Sessions' : 'All Sessions'}</option>
              {SESSIONS.map(session => (
                <option key={session} value={session}>{session}</option>
              ))}
            </select>
         </div>
         <div className="w-full overflow-x-auto custom-scrollbar max-h-[400px]">
            <table className="w-full text-left text-sm min-w-[600px] whitespace-nowrap">
               <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-xs font-semibold sticky top-0 z-10">
                  <tr><th className="p-3 sm:p-4 text-center w-12 text-gray-400">No.</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Status Peringkat' : 'Rank Status'}</th><th className="p-3 sm:p-4">{language === 'id' ? 'Nama Siswa' : 'Student Name'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Sesi' : 'Session'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Tingkat/Level' : 'Tier/Level'}</th><th className="p-3 sm:p-4 text-right">{language === 'id' ? 'Total EXP' : 'Total EXP'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Aksi' : 'Action'}</th></tr>
               </thead>
               <tbody className="divide-y divide-gray-800">
                  {db.students.filter(s => s.status === 'Active' || s.active === 'Active')
                    .filter(s => leaderboardFilter === 'All' ? true : getStudentSession(s) === leaderboardFilter)
                    .map(s => ({ ...s, exp: calculateStudentEXP(s.id, db) }))
                    .sort((a,b) => b.exp - a.exp)
                    .map((s, idx) => {
                       const lvl = getLevelInfo(s.exp);
                       const badge = getLeaderboardBadge(idx);
                       const BadgeIcon = badge.icon;
                       return (
                         <tr key={s.id} className="hover:bg-[#0A0E17] transition-colors">
                            <td className="p-3 sm:p-4 text-center text-gray-500 font-medium">{idx + 1}</td>
                            <td className="p-3 sm:p-4 text-center">
                               <div className={`inline-flex items-center justify-center gap-1.5 px-2.5 py-1 rounded-lg border border-white/5 shadow-sm min-w-[120px] ${badge.bg} ${badge.color}`}>
                                   <BadgeIcon size={14} className={badge.color} />
                                   <span className="font-bold text-xs whitespace-nowrap">{badge.title}</span>
                               </div>
                            </td>
                            <td className="p-3 sm:p-4 font-bold text-white"><div className="flex items-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /> <span className="text-xs text-gray-500 font-normal ml-1">({s.class})</span></div></td>
                            <td className="p-3 sm:p-4 text-center text-gray-300 text-xs">{getStudentSession(s)}</td>
                            <td className="p-3 sm:p-4 text-center"><span className={`px-3 py-1 rounded-full text-[11px] font-semibold uppercase tracking-widest border ${lvl.bg} ${lvl.border} ${lvl.textCol}`}>{lvl.title}</span></td>
                            <td className="p-3 sm:p-4 text-right">
                               <div className="font-black text-[#00D4FF]">{s.exp.toLocaleString()} <span className="text-[11px] text-gray-500 font-bold">EXP</span></div>
                               {expModalStudent?.id === s.id && (
                                  <div className="text-[11px] text-yellow-500/80 font-bold mt-0.5">{language === 'id' ? 'Bonus Saat Ini:' : 'Current Bonus:'} {s.bonusExp || 0}</div>
                               )}
                            </td>
                            <td className="p-3 sm:p-4 text-center">
                               {expModalStudent?.id === s.id ? (
                                  <div className="flex items-center justify-center gap-1.5 animation-fade-in">
                                     <input 
                                        type="number" 
                                        value={expInput} 
                                        onChange={(e) => setExpInput(e.target.value)} 
                                        onKeyDown={(e) => {
                                           if (e.key === 'Enter') {
                                              const val = Number(e.currentTarget.value) || 0;
                                              setDb(prev => ({ ...prev, students: prev.students.map(stu => stu.id === expModalStudent.id ? { ...stu, bonusExp: (Number(stu.bonusExp) || 0) + val, updatedAt: getSyncTimestamp() } : stu) }));
                                              showToast(language === 'id' ? `Berhasil menyesuaikan ${val} EXP untuk ${expModalStudent.name}` : `Successfully adjusted ${val} EXP for ${expModalStudent.name}`);
                                              setExpModalStudent(null);
                                              setExpInput('');
                                           }
                                        }}
                                        placeholder="+/-" 
                                        className="w-14 sm:w-16 bg-[#0B0F19] border border-yellow-500/50 text-white rounded px-2 py-1 text-xs focus:outline-none focus:border-yellow-400 text-center" 
                                        autoFocus 
                                     />
                                     <button type="button" onClick={() => {
                                        const val = Number(expInput) || 0;
                                        setDb(prev => ({ ...prev, students: prev.students.map(stu => stu.id === expModalStudent.id ? { ...stu, bonusExp: (Number(stu.bonusExp) || 0) + val, updatedAt: getSyncTimestamp() } : stu) }));
                                        showToast(language === 'id' ? `Berhasil menyesuaikan ${val} EXP untuk ${expModalStudent.name}` : `Successfully adjusted ${val} EXP for ${expModalStudent.name}`);
                                        setExpModalStudent(null);
                                        setExpInput('');
                                     }} className="p-1 bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 rounded transition-colors" title={language === 'id' ? 'Simpan' : 'Save'}><Check size={14}/></button>
                                     <button type="button" onClick={() => { setExpModalStudent(null); setExpInput(''); }} className="p-1 bg-red-500/10 text-red-400 hover:bg-red-500/20 rounded transition-colors" title={language === 'id' ? 'Cancel' : 'Cancel'}><X size={14}/></button>
                                  </div>
                               ) : (
                                  <button onClick={() => { setExpModalStudent(s); setExpInput(''); }} className="p-1.5 bg-yellow-500/10 text-yellow-400 hover:bg-yellow-500/20 rounded-lg transition-colors border border-yellow-500/20 shadow-sm" title={language === 'id' ? 'Edit EXP' : 'Edit EXP'}>
                                     <Zap size={16} />
                                  </button>
                               )}
                            </td>
                         </tr>
                       )
                    })
                  }
               </tbody>
            </table>
         </div>
      </Card>

    </div>
  );
};

const TutorDashboard = ({ db, setDb, user, setActiveTab, today, isCloudConnected, language = 'en', showToast }: any) => {
  const [expModalStudent, setExpModalStudent] = useState(null);
  const [expInput, setExpInput] = useState('');
  
  const dObj = new Date();
  const currentMonth = String(dObj.getMonth() + 1);
  const currentYear = String(dObj.getFullYear());
  const monthPrefix = `${currentYear}-${currentMonth.padStart(2, '0')}`;

  const mySessions = parseSessions(user.teachingSession);
  const activeStudents = db.students.filter(s => s.status === 'Active');
  const myStudents = activeStudents.filter(s => mySessions.includes(getStudentSession(s))).length;

  // Helper for Co-Teaching: Memisahkan string "Tutor A & Tutor B" untuk mencocokkan nama
  const isMyClass = (tutorString, myName) => tutorString && tutorString.split(' & ').includes(myName);

  // Fix #11: kecualikan event yang dibatalkan/libur dari hitungan kelas hari ini
  // agar badge "Jurnal Tertunda" tidak muncul untuk kelas yang tidak jadi berlangsung.
  const CANCELLED_TYPES = ['Cancelled', 'Holiday', 'Off Day', 'Libur', 'Dibatalkan'];
  const myClassesToday = db.calendar.filter(
    c => c.date === today && isMyClass(c.tutor, user.name) && !CANCELLED_TYPES.includes(c.type)
  );
  const todayClassesCount = myClassesToday.length;

  const myUpcomingClasses = db.calendar
      .filter(c => (c.date || '') >= today && isMyClass(c.tutor, user.name))
      .sort((a,b) => String(a.date || '').localeCompare(String(b.date || '')))
      .slice(0, 5);

  // Fix #4: fallback ke name untuk data legacy yang belum punya tutorId
  const hasCheckedIn = db.tutorAttendance.some(
    a => (a.tutorId === user.id || (!a.tutorId && a.name === user.name)) && a.date === today && a.status === 'Present'
  );
  const checkInText = hasCheckedIn ? (language === 'id' ? 'Hadir' : 'Present') : (language === 'id' ? 'Belum Absen' : 'Not Checked In');

  const journalsToday = db.journals.filter(j => j.tutorName === user.name && j.date === today).length;
  const pendingJournals = Math.max(0, todayClassesCount - journalsToday);

  const studentsInSessionIds = activeStudents.filter(s => mySessions.includes(getStudentSession(s))).map(s => s.id);
  // Fix 4: Also validate the attendance session group matches the student's current session
  const attThisMonth = db.studentAttendance.filter(a => {
    if (!a.date.startsWith(monthPrefix) || !studentsInSessionIds.includes(a.studentId)) return false;
    const studentObj = db.students.find(s => s.id === a.studentId);
    return isAttendanceValidForStudent(a, studentObj);
  });
  const attPresent = attThisMonth.filter(a => a.status === 'Present').length;
  const attendanceRate = attThisMonth.length > 0 ? Math.round((attPresent / attThisMonth.length) * 100) : 0;

  const assessmentsDone = db.assessments.filter(a => Number(a.month) === Number(currentMonth) && String(a.year) === String(currentYear) && mySessions.includes(a.sessionGroup)).length;
  const assessmentsPending = Math.max(0, myStudents - assessmentsDone);

  const classesCompletedMonth = db.calendar.filter(c => isMyClass(c.tutor, user.name) && c.date.startsWith(monthPrefix) && c.date <= today).length;
  // REPLACE THIS SECTION: Mengubah pengali durasi kelas menjadi 1 jam
  const teachingHours = classesCompletedMonth * 1;

  // PERBAIKAN: Hanya tampilkan progress bar untuk kelas yang diajarkan oleh tutor ini saja
  const classProgressData = (mySessions.length > 0 ? mySessions : [user.teachingSession]).map(session => {
     const studentsInSess = db.students.filter(s => s.status === 'Active' && getStudentSession(s) === session).length;
     // Fix #10: hitung unique studentId saja (cegah duplicate assessment inflating progress > 100%)
     const uniqueAssessed = new Set(
       db.assessments
         .filter(a => a.sessionGroup === session && Number(a.month) === Number(currentMonth) && String(a.year) === String(currentYear))
         .map(a => a.studentId)
     ).size;
     const pct = studentsInSess > 0 ? Math.min(100, Math.round((uniqueAssessed / studentsInSess) * 100)) : 0;
     return { session, pct };
  });

  return (
    <div className="space-y-4 sm:space-y-6 animation-fade-in w-full max-w-full overflow-hidden font-sans">
      <GreetingCard userName={user?.name} isCloudConnected={isCloudConnected} language={language} />

      {/* First Row 4 Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3 sm:gap-4">
         <KPICard title={language === 'id' ? 'Siswa Saya' : 'My Students'} value={myStudents} subtext={language === 'id' ? 'Total ditugaskan' : 'Total assigned to you'} icon={Users} colorClass="bg-blue-500" onClick={() => setActiveTab('students')} />
         <KPICard title={language === 'id' ? 'Kelas Hari Ini' : "Today's Classes"} value={todayClassesCount} subtext={language === 'id' ? 'Jadwal hari ini' : 'Scheduled today'} icon={CalendarIcon} colorClass="bg-purple-500" onClick={() => setActiveTab('calendar')} />
         <KPICard title={language === 'id' ? 'Status Absen' : 'Check-in Status'} value={checkInText} subtext={language === 'id' ? 'Kehadiran harian' : 'Daily attendance'} icon={CheckSquare} colorClass={hasCheckedIn ? "bg-emerald-500" : "bg-rose-500"} onClick={() => setActiveTab('tutor_attendance')} />
         <KPICard title={language === 'id' ? 'Jurnal Tertunda' : 'Pending Journals'} value={pendingJournals} subtext={language === 'id' ? 'Belum dikumpulkan' : 'Not yet submitted'} icon={BookOpen} colorClass={pendingJournals === 0 ? "bg-emerald-500" : "bg-amber-500"} onClick={() => setActiveTab('journals')} />
      </div>

      {/* Second Row 4 Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3 sm:gap-4">
         <KPICard title={language === 'id' ? 'Tingkat Kehadiran' : 'Attendance Rate'} value={`${attendanceRate}%`} subtext={language === 'id' ? 'Rata-rata sesi' : 'Session average'} icon={BarChart3} colorClass="bg-teal-500" onClick={() => setActiveTab('student_attendance')} />
         <KPICard title={language === 'id' ? 'Penilaian Tertunda' : 'Assessments Pending'} value={assessmentsPending} subtext={language === 'id' ? 'Sisa bulan ini' : 'Monthly remaining'} icon={FileText} colorClass={assessmentsPending === 0 ? "bg-emerald-500" : "bg-rose-500"} onClick={() => setActiveTab('assessments')} />
         <KPICard title={language === 'id' ? 'Jam Mengajar' : 'Teaching Hours'} value={`${teachingHours}h`} subtext={language === 'id' ? 'Bulan ini' : 'This month'} icon={Clock} colorClass="bg-indigo-500" />
         <KPICard title={language === 'id' ? 'Kelas Selesai' : 'Classes Completed'} value={classesCompletedMonth} subtext={language === 'id' ? 'Bulan ini' : 'This month'} icon={GraduationCap} colorClass="bg-blue-500" />
      </div>

      {/* 2 Columns: Today's Schedule & Tutor Tasks */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-6 w-full">
         <Card className="border-t-4 border-t-[#00D4FF] shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><CalendarIcon size={18} className="text-[#00D4FF]"/> {language === 'id' ? 'Jadwal Berikutnya' : 'Next Schedule'}</h3>
               <Button variant="ghost" className="text-xs py-2 px-3 min-h-[36px]" onClick={() => setActiveTab('calendar')}>{language === 'id' ? 'Kalender Penuh' : 'Full Calendar'}</Button>
            </div>
            <div className="p-4 sm:p-5 space-y-3 bg-[#151B26] flex-1">
               {myUpcomingClasses.map(c => (
                  <div key={c.id} className="flex justify-between items-center p-3 sm:p-4 bg-[#0B0F19] rounded-xl border border-gray-800 border-l-4 border-l-[#00D4FF] hover:bg-[#0A0E17] transition-colors gap-3">
                     <div className="flex-1 min-w-0">
                        <p className="font-bold text-white text-sm sm:text-base truncate">{c.sessionGroup || c.name}</p>
                        <div className="flex items-center flex-wrap gap-2 mt-1 sm:mt-1.5 text-xs">
                           <span className="text-[#00D4FF] font-medium whitespace-nowrap">
                             <CalendarIcon size={12} className="inline mr-1" />
                             {safeDateDisplay(c.date, language === 'id' ? 'id-ID' : 'en-US', {weekday: 'short'})}, {c.date}
                           </span>
                           <span className="text-gray-400 font-medium whitespace-nowrap"><Clock size={12} className="inline mr-1" />{c.startTime} - {c.endTime}</span>
                           <span className="text-amber-400 font-bold bg-amber-400/10 px-2 py-0.5 rounded ml-auto">{calculateDaysLeft(c.date, today)}</span>
                        </div>
                     </div>
                  </div>
               ))}
               {myUpcomingClasses.length === 0 && <p className="text-xs sm:text-sm text-gray-500 text-center py-4">{language === 'id' ? 'Tidak ada jadwal kelas mendatang.' : 'No upcoming classes scheduled.'}</p>}
            </div>
         </Card>

         <Card className="border-t-4 border-t-purple-500 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><CheckCircle2 size={18} className="text-purple-400"/> {language === 'id' ? 'Tugas Tutor' : 'Tutor Tasks'}</h3>
            </div>
            <div className="p-4 sm:p-5 space-y-4 bg-[#151B26] flex-1 flex flex-col justify-center">
               <Button onClick={() => setActiveTab('journals')} className="w-full justify-start py-4 text-left border border-gray-700 bg-[#0B0F19] hover:bg-[#1A2234] hover:border-purple-500/50 transition-all text-white shadow-md rounded-xl" variant="secondary">
                 <BookOpen size={20} className="text-purple-400 mr-2" />
                 <div className="flex flex-col items-start">
                   <span className="font-bold">{language === 'id' ? 'Isi Jurnal Belajar' : 'Fill Learning Journal'}</span>
                   <span className="text-xs text-gray-400 font-normal mt-0.5">{language === 'id' ? 'Catat materi hari ini' : "Record today's teaching materials"}</span>
                 </div>
               </Button>
               <Button onClick={() => setActiveTab('student_attendance')} className="w-full justify-start py-4 text-left border border-gray-700 bg-[#0B0F19] hover:bg-[#1A2234] hover:border-teal-500/50 transition-all text-white shadow-md rounded-xl" variant="secondary">
                 <UserCheck size={20} className="text-teal-400 mr-2" />
                 <div className="flex flex-col items-start">
                   <span className="font-bold">{language === 'id' ? 'Isi Kehadiran' : 'Complete Attendance'}</span>
                   <span className="text-xs text-gray-400 font-normal mt-0.5">{language === 'id' ? 'Fields markedi kehadiran siswa' : "Mark students' presence"}</span>
                 </div>
               </Button>
               <Button onClick={() => setActiveTab('assessments')} className="w-full justify-start py-4 text-left border border-gray-700 bg-[#0B0F19] hover:bg-[#1A2234] hover:border-rose-500/50 transition-all text-white shadow-md rounded-xl" variant="secondary">
                 <FileText size={20} className="text-rose-400 mr-2" />
                 <div className="flex flex-col items-start">
                   <span className="font-bold">{language === 'id' ? 'Kirim Penilaian' : 'Submit Assessment'}</span>
                   <span className="text-xs text-gray-400 font-normal mt-0.5">{language === 'id' ? 'Masukkan nilai bulanan' : 'Input monthly grades'}</span>
                 </div>
               </Button>
            </div>
         </Card>
      </div>

      {/* Bottom Section: Class Progress & Announcements */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-6 w-full">
         <Card className="border-t-4 border-t-emerald-400 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><BarChart3 size={18} className="text-emerald-400"/> {language === 'id' ? 'Progres Kelas' : 'Class Progress'}</h3>
               <span className="text-xs text-gray-400">{language === 'id' ? 'Penilaian selesai' : 'Assessments completed'}</span>
            </div>
            <div className="p-4 sm:p-6 space-y-6 bg-[#151B26] flex-1">
               {classProgressData.map((cls, idx) => {
                 const colors = ['bg-blue-500', 'bg-purple-500', 'bg-emerald-500', 'bg-amber-500', 'bg-rose-500'];
                 const color = colors[idx % colors.length];
                 return (
                   <div key={cls.session}>
                      <div className="flex justify-between items-end mb-2">
                         <span className="text-gray-300 font-semibold text-sm">{cls.session.replace(' Session', '')}</span>
                         <span className="text-white font-black text-sm leading-none">{cls.pct}%</span>
                      </div>
                      <div className="w-full bg-gray-800/80 rounded-full h-2.5 overflow-hidden shadow-inner">
                         <div className={`h-full rounded-full transition-all duration-1000 ${color}`} style={{ width: `${cls.pct}%` }}></div>
                      </div>
                   </div>
                 )
               })}
            </div>
         </Card>

         <Card className="border-t-4 border-t-yellow-400 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-gray-800 bg-[#151B26]">
               <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><Bell size={18} className="text-yellow-400"/> {language === 'id' ? 'Pengumuman Terbaru' : 'Recent Announcements'}</h3>
               <Button variant="ghost" className="text-xs py-2 px-3 min-h-[36px]" onClick={() => setActiveTab('announcements')}>{language === 'id' ? 'Lihat Semua' : 'View All'}</Button>
            </div>
            <div className="p-4 sm:p-5 space-y-3 sm:space-y-4 bg-[#151B26] flex-1 overflow-y-auto max-h-[300px] custom-scrollbar">
               {db.announcements.slice(-3).reverse().map(a => (
                  <div key={a.id} className="p-3 sm:p-4 bg-[#0B0F19] rounded-xl border border-gray-800 hover:border-gray-700 transition-colors">
                     <h4 className="font-bold text-[#00D4FF] text-sm sm:text-base mb-1">{a.title}</h4>
                     <p className="text-xs text-gray-500 mb-1.5 sm:mb-2 font-medium">{a.date}</p>
                     <p className="text-xs sm:text-sm text-gray-300 line-clamp-2">{a.content}</p>
                  </div>
               ))}
               {db.announcements.length === 0 && <p className="text-xs sm:text-sm text-gray-500 text-center py-4">{language === 'id' ? 'Tidak ada pengumuman.' : 'No announcements.'}</p>}
            </div>
         </Card>
      </div>

      {/* NEW: CLASS LEADERBOARD SECTION (TUTOR ONLY) */}
      <Card className="p-0 overflow-hidden border-t-4 border-t-yellow-400 shadow-lg mt-4 sm:mt-6">
         <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex justify-between items-center">
            <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2"><Trophy size={20} className="text-yellow-400" /> {language === 'id' ? 'Peringkat Kelas' : 'Class Leaderboard'}</h3>
            <div className="flex flex-wrap gap-1">{mySessions.map(s => <span key={s} className="text-xs text-gray-400 bg-gray-800 px-3 py-1 rounded-full border border-gray-700">{s}</span>)}</div>
         </div>
         <div className="w-full overflow-x-auto custom-scrollbar max-h-[350px]">
            <table className="w-full text-left text-sm min-w-[500px] whitespace-nowrap">
               <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-xs font-semibold sticky top-0 z-10">
                  <tr><th className="p-3 sm:p-4 text-center w-12 text-gray-400">No.</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Status Peringkat' : 'Rank Status'}</th><th className="p-3 sm:p-4">{language === 'id' ? 'Nama Siswa' : 'Student Name'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Tingkat/Level' : 'Tier/Level'}</th><th className="p-3 sm:p-4 text-right">{language === 'id' ? 'Total EXP' : 'Total EXP'}</th><th className="p-3 sm:p-4 text-center">{language === 'id' ? 'Aksi' : 'Action'}</th></tr>
               </thead>
               <tbody className="divide-y divide-gray-800">
                  {activeStudents.filter(s => mySessions.includes(getStudentSession(s)))
                    .map(s => ({ ...s, exp: calculateStudentEXP(s.id, db) }))
                    .sort((a,b) => b.exp - a.exp)
                    .map((s, idx) => {
                       const lvl = getLevelInfo(s.exp);
                       const badge = getLeaderboardBadge(idx);
                       const BadgeIcon = badge.icon;
                       return (
                         <tr key={s.id} className="hover:bg-[#0A0E17] transition-colors">
                            <td className="p-3 sm:p-4 text-center text-gray-500 font-medium">{idx + 1}</td>
                            <td className="p-3 sm:p-4 text-center">
                               <div className={`inline-flex items-center justify-center gap-1.5 px-2.5 py-1 rounded-lg border border-white/5 shadow-sm min-w-[120px] ${badge.bg} ${badge.color}`}>
                                   <BadgeIcon size={14} className={badge.color} />
                                   <span className="font-bold text-xs whitespace-nowrap">{badge.title}</span>
                               </div>
                            </td>
                            <td className="p-3 sm:p-4 font-bold text-white"><div className="flex items-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /> <span className="text-xs text-gray-500 font-normal ml-1">({s.class})</span></div></td>
                            <td className="p-3 sm:p-4 text-center"><span className={`px-3 py-1 rounded-full text-[11px] font-semibold uppercase tracking-widest border ${lvl.bg} ${lvl.border} ${lvl.textCol}`}>{lvl.title}</span></td>
                            <td className="p-3 sm:p-4 text-right">
                               <div className="font-black text-[#00D4FF]">{s.exp.toLocaleString()} <span className="text-[11px] text-gray-500 font-bold">EXP</span></div>
                               {expModalStudent?.id === s.id && (
                                  <div className="text-[11px] text-yellow-500/80 font-bold mt-0.5">{language === 'id' ? 'Bonus Saat Ini:' : 'Current Bonus:'} {s.bonusExp || 0}</div>
                               )}
                            </td>
                            <td className="p-3 sm:p-4 text-center">
                               {expModalStudent?.id === s.id ? (
                                  <div className="flex items-center justify-center gap-1.5 animation-fade-in">
                                     <input 
                                        type="number" 
                                        value={expInput} 
                                        onChange={(e) => setExpInput(e.target.value)} 
                                        onKeyDown={(e) => {
                                           if (e.key === 'Enter') {
                                              const val = Number(e.currentTarget.value) || 0;
                                              setDb(prev => ({ ...prev, students: prev.students.map(stu => stu.id === expModalStudent.id ? { ...stu, bonusExp: (Number(stu.bonusExp) || 0) + val, updatedAt: getSyncTimestamp() } : stu) }));
                                              showToast(language === 'id' ? `Berhasil menyesuaikan ${val} EXP untuk ${expModalStudent.name}` : `Successfully adjusted ${val} EXP for ${expModalStudent.name}`);
                                              setExpModalStudent(null);
                                              setExpInput('');
                                           }
                                        }}
                                        placeholder="+/-" 
                                        className="w-14 sm:w-16 bg-[#0B0F19] border border-yellow-500/50 text-white rounded px-2 py-1 text-xs focus:outline-none focus:border-yellow-400 text-center" 
                                        autoFocus 
                                     />
                                     <button type="button" onClick={() => {
                                        const val = Number(expInput) || 0;
                                        setDb(prev => ({ ...prev, students: prev.students.map(stu => stu.id === expModalStudent.id ? { ...stu, bonusExp: (Number(stu.bonusExp) || 0) + val, updatedAt: getSyncTimestamp() } : stu) }));
                                        showToast(language === 'id' ? `Berhasil menyesuaikan ${val} EXP untuk ${expModalStudent.name}` : `Successfully adjusted ${val} EXP for ${expModalStudent.name}`);
                                        setExpModalStudent(null);
                                        setExpInput('');
                                     }} className="p-1 bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 rounded transition-colors" title={language === 'id' ? 'Simpan' : 'Save'}><Check size={14}/></button>
                                     <button type="button" onClick={() => { setExpModalStudent(null); setExpInput(''); }} className="p-1 bg-red-500/10 text-red-400 hover:bg-red-500/20 rounded transition-colors" title={language === 'id' ? 'Cancel' : 'Cancel'}><X size={14}/></button>
                                  </div>
                               ) : (
                                  <button onClick={() => { setExpModalStudent(s); setExpInput(''); }} className="p-1.5 bg-yellow-500/10 text-yellow-400 hover:bg-yellow-500/20 rounded-lg transition-colors border border-yellow-500/20 shadow-sm" title={language === 'id' ? 'Edit EXP' : 'Edit EXP'}>
                                     <Zap size={16} />
                                  </button>
                               )}
                            </td>
                         </tr>
                       )
                    })
                  }
               </tbody>
            </table>
         </div>
      </Card>

    </div>
  );
};

const Dashboard = ({ db, setDb, user, setActiveTab, isCloudConnected, language, showToast }: any) => {
  const today = getTodayDateLocal();
  if (user.role === 'student') {
    return <StudentDashboard db={db} user={user} setActiveTab={setActiveTab} today={today} isCloudConnected={isCloudConnected} language={language} />;
  }
  if (user.role === 'tutor') {
    return <TutorDashboard db={db} setDb={setDb} user={user} setActiveTab={setActiveTab} today={today} isCloudConnected={isCloudConnected} language={language} showToast={showToast} />;
  }
  return <AdminDashboard db={db} setDb={setDb} user={user} setActiveTab={setActiveTab} today={today} isCloudConnected={isCloudConnected} language={language} showToast={showToast} />;
};

class ErrorBoundary extends React.Component<any, any> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError(error: any) {
    return { hasError: true };
  }
  componentDidCatch(error: any, errorInfo: any) {
    console.error("Global UI Crash Intercepted:", error, errorInfo);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-[#0B0F19] flex items-center justify-center p-4 font-sans text-white">
          <div className="max-w-md w-full bg-[#151B26] border border-red-500/30 p-8 rounded-2xl shadow-2xl text-center">
            <div className="w-16 h-16 bg-red-500/10 rounded-full flex items-center justify-center mx-auto mb-4 border border-red-500/20">
              <AlertCircle size={32} className="text-red-500" />
            </div>
            <h2 className="text-2xl font-bold text-white mb-2">System is Recovering</h2>
            <p className="text-gray-400 text-sm mb-6">An unexpected data anomaly occurred. Don't panic — just reload the page to restore your session.</p>
            <div className="flex flex-col gap-3">
              <button onClick={() => window.location.reload()} className="bg-red-500 hover:bg-red-600 text-white font-bold py-3 px-6 rounded-xl transition-all shadow-[0_0_15px_rgba(239,68,68,0.3)]">
                Reload Application
              </button>
              {/* PERBAIKAN: Tombol Hard Reset untuk mencegah Infinite Crash Loop jika memori lokal korup */}
              <button onClick={() => { localStorage.removeItem('ecg_db'); window.location.reload(); }} className="text-xs text-gray-500 hover:text-gray-300 underline transition-colors">
                Still failing? Click to Hard Reset (Clear Cache)
              </button>
            </div>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// PERBAIKAN: Buat struktur default untuk memastikan tidak ada array yang 'undefined'
const defaultDbStructure = {
  users: [], students: [], tutors: [], studentAttendance: [], tutorAttendance: [],
  journals: [], assessments: [], payments: [], payroll: [], calendar: [], announcements: [], recycleBin: [], materials: []
};

// Fungsi untuk menormalisasi data lama/eksternal yang menggunakan key bahasa Indonesia
const normalizeData = (data) => {
   // FIX (Silent Overwrite Race Condition): Pisahkan logs dari db utama agar
   // perubahan log tidak memicu sync useEffect. Logs disimpan ke state terpisah.
   const { auditLogs: _a, debugLogs: _d, ...dataWithoutLogs } = data || {};
   const norm = { ...defaultDbStructure, ...dataWithoutLogs };
   
   // 🛡️ STRICT ARRAY ENFORCEMENT (Data Type Guardian)
   Object.keys(defaultDbStructure).forEach(key => {
      if (!Array.isArray(norm[key])) {
         norm[key] = [];
      }
   });

   // --- PATCH BUG: FIX TANGGAL & WAKTU ANEH DARI GOOGLE SHEETS (LMT BATAVIA BUG) ---
   // BUGFIX #1 & #5: Gunakan metode UTC (getUTCHours, getUTCFullYear, dst.)
   // agar tidak terjadi pergeseran jam/tanggal akibat konversi timezone browser.
   // Google Sheets menyimpan time/date sebagai UTC — membacanya dengan getHours()
   // (local time) menyebabkan geser ±7 jam di WIB.
   const fixTime = (t) => {
      if (typeof t === 'string' && t.includes('T')) {
         if (t.includes('Z')) {
            // Logika lama untuk menangani sisa cache format UTC 
            const d = new Date(t);
            if (!isNaN(d.getTime())) {
               let hours = d.getUTCHours();
               let minutes = d.getUTCMinutes();
               let seconds = d.getUTCSeconds();
               // Kompensasi bug zona waktu Batavia (LMT +07:07:12) sebelum tahun 1900
               if (d.getUTCFullYear() <= 1901 && (seconds !== 0 || minutes % 5 !== 0)) {
                  minutes += 7;
                  seconds += 12;
                  if (seconds >= 60) { seconds -= 60; minutes += 1; }
                  if (minutes >= 60) { minutes -= 60; hours += 1; }
               }
               minutes = Math.round(minutes / 5) * 5;
               if (minutes >= 60) { minutes = 0; hours += 1; }
               hours = hours % 24;
               return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
            }
         } else {
            // Logika baru: Ekstrak HH:mm langsung dari format ISO ber-offset (+07:00)
            const timePart = t.split('T')[1];
            if (timePart) {
               return timePart.substring(0, 5);
            }
         }
      }
      return t;
   };
   const fixDate = (dStr) => {
      if (typeof dStr === 'string' && dStr.includes('T')) {
         if (dStr.includes('Z')) {
            // Konversi cache format UTC lama
            const d = new Date(dStr);
            if (!isNaN(d.getTime())) {
               return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
            }
         } else {
            // Ambil langsung YYYY-MM-DD dari format offset baru dari Code.gs
            return dStr.split('T')[0];
         }
      }
      return dStr;
   };
   const fixDateTime = (dStr) => {
      if (typeof dStr === 'string' && dStr.includes('T')) {
         // Biarkan Date parsing menangani string offset valid terbaru dengan zona waktu otomatis browser
         const d = new Date(dStr);
         if (!isNaN(d.getTime())) {
            return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}, ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
         }
      }
      return dStr;
   };

   if (norm.calendar && Array.isArray(norm.calendar)) {
      norm.calendar = norm.calendar.map(c => ({
         ...c,
         date: fixDate(c.date),
         startTime: fixTime(c.startTime),
         endTime: fixTime(c.endTime)
      }));
   }

   ['studentAttendance', 'tutorAttendance', 'journals', 'payments', 'payroll', 'announcements', 'materials'].forEach(collection => {
      if (norm[collection] && Array.isArray(norm[collection])) {
         norm[collection] = norm[collection].map(item => ({
            ...item,
            date: item.date ? fixDate(item.date) : item.date,
            time: item.time ? fixTime(item.time) : item.time
         }));
      }
   });

   // Bersihkan juga tanggal bergabung (joinedDate) siswa & tutor dari format ISO/UTC mentah
   if (norm.students && Array.isArray(norm.students)) {
      norm.students = norm.students.map(item => ({
         ...item,
         joinedDate: item.joinedDate ? fixDate(item.joinedDate) : item.joinedDate,
         whatsapp: item.whatsapp ? normalizeWhatsapp(item.whatsapp) : item.whatsapp
      }));
   }
   if (norm.tutors && Array.isArray(norm.tutors)) {
      norm.tutors = norm.tutors.map(item => ({
         ...item,
         joinedDate: item.joinedDate ? fixDate(item.joinedDate) : item.joinedDate,
         phone: item.phone ? normalizeWhatsapp(item.phone) : item.phone
      }));
   }

   // Bersihkan tanggal/waktu bersarang pada tugas (materials): submission siswa & balasan diskusi
   if (norm.materials && Array.isArray(norm.materials)) {
      norm.materials = norm.materials.map(mat => ({
         ...mat,
         submissions: Array.isArray(mat.submissions) ? mat.submissions.map(sub => ({
            ...sub,
            date: sub.date ? fixDateTime(sub.date) : sub.date,
            replies: Array.isArray(sub.replies) ? sub.replies.map(r => ({
               ...r,
               date: r.date ? fixDateTime(r.date) : r.date
            })) : sub.replies
         })) : mat.submissions
      }));
   }

   // CLEANUP RECYCLE BIN: Hapus permanen data yang sudah lebih dari 30 hari di tong sampah
   if (norm.recycleBin && Array.isArray(norm.recycleBin)) {
      const THIRTY_DAYS = 30 * 24 * 60 * 60 * 1000;
      norm.recycleBin = norm.recycleBin.filter(b => {
         if (!b.deletedAt) return true;
         const delTime = new Date(String(b.deletedAt).replace(' ', 'T')).getTime();
         return isNaN(delTime) || (Date.now() - delTime < THIRTY_DAYS);
      });
   }
   // --- END PATCH BUG ---

   if (norm.users && Array.isArray(norm.users)) {
      norm.users = norm.users.map(u => {
         let r = u.role ? String(u.role).toLowerCase() : 'student';
         if (r === 'siswa') r = 'student';
         if (r === 'guru' || r === 'pengajar') r = 'tutor';
         
         return {
            ...u,
            name: u.name || u.nama || 'Unknown',
            role: r,
            // Jika tidak ada username, generate dari nama
            // BUG FIX B: trim() wajib — Google Sheets sering menambah spasi di awal/akhir
            // yang tidak tampak di UI tapi menyebabkan === comparison login selalu gagal.
            username: String(u.username || (u.nama ? String(u.nama).toLowerCase().replace(/[^a-z0-9]/g, '') : `user_${Math.floor(Math.random()*10000)}`)).trim().toLowerCase(),
            password: (u.password !== undefined && u.password !== null && u.password !== '') ? String(u.password).trim() : '',
            active: (u.active || 'Active').toString().trim(),
            mustChangePassword: u.password ? u.mustChangePassword : true
         };
      });
   }
   
   // PERBAIKAN: Pastikan Super Admin SELALU ada di database meskipun data Cloud kosong
   const hasSuperAdmin = norm.users.some(u => (u.username || '').toLowerCase() === 'vicky' || String(u.role).toLowerCase().includes('super'));
   if (!hasSuperAdmin) {
      norm.users.unshift({
         id: 'ADM-FALLBACK',
         username: 'vicky',
         password: 'password',
         role: 'super admin',
         name: 'Vicky',
         active: 'Active',
         mustChangePassword: true  // FIX: Paksa ganti password default saat pertama login
      });
   }
   
   if (norm.tutors && Array.isArray(norm.tutors)) {
      norm.tutors = norm.tutors.map(t => ({
         ...t,
         // BUG FIX A: username/password null/undefined harus jadi string kosong '',
         // bukan tetap null — supaya perbandingan === di login tidak selalu false.
         username: (t.username !== undefined && t.username !== null) ? String(t.username).trim().toLowerCase() : '',
         password: (t.password !== undefined && t.password !== null && t.password !== '') ? String(t.password).trim() : '',
         // BUG FIX A: status tutor harus selalu ada dan bertipe string
         status: t.status || 'Active',
         mustChangePassword: t.mustChangePassword !== undefined ? t.mustChangePassword : false
      }));
   }
   
   return norm;
};

// KOMPONEN BARU: Indikator Auto-Save Real-time
// Status: 'saving' (lokal instan) | 'syncing' (mengirim ke cloud) | 'saved' | 'error'
const CloudAutoSaveIndicator = ({ status, language }: { status: string, language: string }) => {
  if (status === 'saving') {
    // Fase 1: Data sudah aman di localStorage, sedang menunggu kirim ke cloud
    return (
      <div className="flex items-center gap-2 bg-[#0A0E17]/95 backdrop-blur-md border border-blue-500/40 px-3.5 py-1.5 rounded-full text-blue-400 text-xs font-bold shadow-[0_0_15px_rgba(59,130,246,0.2)] animation-fade-in pointer-events-auto">
        <Database size={14} className="animate-pulse" />
        <span className="tracking-wide">{language === 'id' ? 'Tersimpan Lokal...' : 'Saved Locally...'}</span>
      </div>
    );
  }
  if (status === 'syncing') {
    // Fase 2: Request sedang dikirim ke Google Apps Script
    return (
      <div className="flex items-center gap-2 bg-[#0A0E17]/95 backdrop-blur-md border border-yellow-500/40 px-3.5 py-1.5 rounded-full text-yellow-400 text-xs font-bold shadow-[0_0_15px_rgba(234,179,8,0.2)] animation-fade-in pointer-events-auto">
        <RefreshCw size={14} className="animate-spin" />
        <span className="tracking-wide">{language === 'id' ? 'Mengirim ke Cloud...' : 'Sending to Cloud...'}</span>
      </div>
    );
  }
  if (status === 'saved') {
    return (
      <div className="flex items-center gap-2 bg-[#0A0E17]/95 backdrop-blur-md border border-emerald-500/40 px-3.5 py-1.5 rounded-full text-emerald-400 text-xs font-bold shadow-[0_0_15px_rgba(16,185,129,0.2)] animation-fade-in pointer-events-auto transition-opacity duration-500">
        <Cloud size={14} />
        <span className="tracking-wide">{language === 'id' ? 'Tersimpan di Cloud' : 'Saved to Cloud'}</span>
      </div>
    );
  }
  if (status === 'error') {
    return (
      <div className="flex items-center gap-2 bg-[#0A0E17]/95 backdrop-blur-md border border-red-500/40 px-3.5 py-1.5 rounded-full text-red-400 text-xs font-bold shadow-[0_0_15px_rgba(239,68,68,0.2)] animation-fade-in pointer-events-auto">
        <CloudOff size={14} />
        <span className="tracking-wide">{language === 'id' ? 'Gagal Simpan' : 'Sync Error'}</span>
      </div>
    );
  }
  return null;
};

function MainApp() {
  const [db, setDbState] = useState(defaultDbStructure);

  // Semua mutation dari module menggunakan wrapper ini. Jika sebuah record berubah
  // tetapi caller lupa memberi updatedAt, wrapper otomatis menambahkan timestamp UTC.
  // Jalur cloud/load memakai setDbState langsung agar pull dari server tidak dianggap edit lokal.
  const setDb = (updater: any) => {
    setDbState((prevDb: any) => {
      const nextDb = typeof updater === 'function' ? updater(prevDb) : updater;
      if (!nextDb || typeof nextDb !== 'object') return nextDb;

      const now = getSyncTimestamp();
      const stamped: any = { ...nextDb };
      SYNC_COLLECTIONS.forEach(col => {
        const previous = Array.isArray(prevDb?.[col]) ? prevDb[col] : [];
        const next = Array.isArray(nextDb?.[col]) ? nextDb[col] : [];
        const previousById = new Map<string, any>();
        previous.forEach(item => {
          const id = getRecordId(item);
          if (id) previousById.set(id, item);
        });
        stamped[col] = next.map(item => {
          const id = getRecordId(item);
          if (!id || col === 'recycleBin') return item;
          const oldItem = previousById.get(id);
          const changed = !oldItem || !syncRecordsEqual(oldItem, item);
          if (!changed) return item;
          // Setiap mutation lokal harus mendapatkan timestamp baru, termasuk saat
          // caller me-spread record lama yang sudah memiliki updatedAt.
          return { ...item, updatedAt: now };
        });
      });
      return stamped;
    });
  };
  
  const latestDbRef = useRef(db);
  useEffect(() => { latestDbRef.current = db; }, [db]);

  // FIX (Silent Overwrite Race Condition): Logs dipisahkan dari db utama agar
  // penambahan audit/debug log TIDAK memicu sync useEffect yang akan menimpa
  // data cloud dengan state lokal yang belum tentu ter-update.
  const [logs, setLogs] = useState({ auditLogs: [], debugLogs: [] });

  const [savedAccounts, setSavedAccounts] = useState(() => {
    // BUG FIX I: Error handling untuk localStorage korup.
    // Jika data truncated/invalid, app tidak boleh crash — reset ke array kosong.
    // Sekaligus migrasi username lama (uppercase) ke lowercase agar konsisten dengan fix F.
    try {
      const saved = localStorage.getItem('ecg_saved_accounts');
      if (!saved) return [];
      const parsed = JSON.parse(saved);
      if (!Array.isArray(parsed)) return [];
      // Migrasi: normalisasi semua username ke lowercase & hapus duplikat
      const seen = new Set();
      return parsed
        .filter(acc => acc && typeof acc.username === 'string')
        .map(acc => ({ ...acc, username: acc.username.trim().toLowerCase() }))
        .filter(acc => {
          if (seen.has(acc.username)) return false;
          seen.add(acc.username);
          return true;
        });
    } catch (e) {
      localStorage.removeItem('ecg_saved_accounts');
      return [];
    }
  });
  const [showAccountMenu, setShowAccountMenu] = useState(false);

  const [currentUser, setCurrentUser] = useState(() => {
    const saved = sessionStorage.getItem('ecg_active_session');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Normalisasi role agar tidak terjadi blank page jika user adalah super admin
        if (String(parsed.role).toLowerCase().includes('super')) {
           parsed.role = 'admin';
           parsed.isSuperAdmin = true;
        }
        return parsed;
      } catch (e) {
        sessionStorage.removeItem('ecg_active_session');
        return null;
      }
    }
    return null;
  });
  
  const [activeTab, setActiveTab] = useState(() => {
    const hash = window.location.hash.replace('#', '').split('?')[0];
    return hash && hash !== 'dashboard' ? hash : 'dashboard';
  });
  
  const [isSidebarOpen, setSidebarOpen] = useState(true);

  useEffect(() => {
    if (currentUser && !currentUser.mustChangePassword) {
      setSidebarOpen(true);
    }
  }, [currentUser?.mustChangePassword]);
  const [toast, setToast] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [logoutConfirm, setLogoutConfirm] = useState(false);
  const [isDbLoaded, setIsDbLoaded] = useState(false);
  
  // State untuk Cloud Connection
  const [isCloudConnected, setIsCloudConnected] = useState(true);
  const [syncStatus, setSyncStatus] = useState('saved'); // State Baru: 'saved' | 'syncing' | 'error'
  const prevCloudState = useRef(true);

  // State untuk Bahasa (hanya aktif untuk student)
  const [language, setLanguage] = useState(() => localStorage.getItem('ecg_lang') || 'en');
  useEffect(() => { localStorage.setItem('ecg_lang', language); }, [language]);

  // FITUR DEWA: PWA (Progressive Web App) Installer
  useEffect(() => {
    // Mewarnai bar notifikasi HP (Status Bar) menjadi warna biru gelap aplikasi
    if (!document.querySelector('meta[name="theme-color"]')) {
       const themeMeta = document.createElement('meta');
       themeMeta.name = 'theme-color';
       themeMeta.content = '#0B0F19';
       document.head.appendChild(themeMeta);
    }

    // Mengaktifkan Service Worker (PWA) secara diam-diam agar tidak ada peringatan merah 404
    // jika file sw.js dan manifest.json belum dibuat di server hosting Anda.
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('/sw.js').catch(() => {
           // Silent fail: Abaikan jika hosting belum memiliki file PWA
        });
      });
    }
  }, []);

  // PERBAIKAN: Ref untuk mencegah data lokal menimpa data cloud saat aplikasi pertama kali dimuat
  const skipCloudSave = useRef(true);

  // PERBAIKAN KRITIS: Ref untuk menandai jika user sudah menginput data baru (Mencegah Race Condition)
  const isDbDirty = useRef(false);

  // Fix 6: Dirty-state guard — modules call setModuleDirty when user starts filling forms
  const moduleDirtyRef = useRef(false);
  const setModuleDirty = (dirty: boolean) => { moduleDirtyRef.current = dirty; };

  // FIX (Silent Overwrite Race Condition): Snapshot entitas intel dari db sebelumnya.
  // Digunakan oleh sync useEffect untuk mendeteksi apakah perubahan db nyata (entitas)
  // atau hanya log/field non-kritis. Jika hanya log, sync ke cloud DILANGWANGI.
  const prevEntitiesRef = useRef(null);

  // FIX: Ref untuk membaca logs terbaru di dalam setTimeout closure sync (menghindari stale state).
  const logsRef = useRef({ auditLogs: [], debugLogs: [] });
  useEffect(() => { logsRef.current = logs; }, [logs]);

  // NEW: Ref untuk melacak versi database (Version Control & Anti-Conflict)
  const dbVersion = useRef(null);

  // NEW: Ref untuk timer double-click tombol back (Exit App)
  const exitToastTimeout = useRef(null);
  // PERBAIKAN: Ref untuk memastikan Toast notifikasi tidak bertabrakan hilangnya
  const toastTimeoutRef = useRef(null);
  // NEW: Ref untuk debounce sync ke cloud (mencegah race condition saat input beruntun)
  const syncDebounceTimer = useRef(null);
  // BUGFIX #6: Ganti window._syncBusyAttempt (global, race condition multi-tab)
  // dengan useRef yang scoped ke instance komponen ini saja.
  const syncBusyAttempt = useRef(0);
  // Satu requestId dipertahankan selama retry payload YANG SAMA agar retry setelah
  // timeout tidak menulis dua kali. Bila user mengubah data sebelum retry, requestId
  // harus diganti supaya cache idempotency server tidak mengembalikan hasil request lama.
  const activeSyncRequestIdRef = useRef(null);
  const activeSyncFingerprintRef = useRef(null);
  const syncRequestCounterRef = useRef(0);
  // FIX DELTA PAYLOAD: Snapshot db terakhir yang BERHASIL tersinkron ke cloud.
  // Digunakan untuk menghitung delta (koleksi mana yang berubah) sebelum kirim ke server.
  // Dengan ini, kita TIDAK mengirim seluruh db — hanya koleksi yang benar-benar berubah,
  // sehingga data koleksi lain yang mungkin diubah user lain tidak akan tertimpa.
  // Baseline sync dipersist ke localStorage setelah pull/sync berhasil.
  // Bila baseline tersedia, delta dihitung record-per-record terhadap baseline tersebut.
  // Bila belum tersedia (misalnya perangkat baru), pull tidak pernah menganggap
  // record cloud yang tidak ada di cache sebagai deletion.
  const lastSyncedSnapshotRef = useRef<string | null>(null);
  // NEW: Ref untuk mencegah infinite loop saat proses exit browser
  const isExiting = useRef(false);

  // BUGFIX: CEGAH TAB DITUTUP SAAT SEDANG SYNC ATAU ADA DATA BELUM TERSIMPAN
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      // Jika database masih kotor (baru diketik) atau sedang proses sinkronisasi
      if (isDbDirty.current || syncStatus === 'syncing' || syncStatus === 'saving') {
        e.preventDefault();
        e.returnValue = 'Data Anda sedang disimpan ke Cloud. Yakin ingin keluar?';
        return e.returnValue;
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [syncStatus]);

  const getSyncIdentity = () => {
    try {
      const active = sessionStorage.getItem('ecg_active_session');
      if (active) {
        const parsed = JSON.parse(active);
        const identity = parsed?.id || parsed?.username || parsed?.name;
        if (identity) return String(identity).trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_');
      }
    } catch (e) {}
    const identity = currentUser?.id || currentUser?.username || currentUser?.name;
    return identity ? String(identity).trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_') : 'anonymous';
  };

  const persistSyncBaseline = () => {
    try {
      const identity = getSyncIdentity();
      if (lastSyncedSnapshotRef.current) localStorage.setItem(`ecg_sync_snapshot:${identity}`, lastSyncedSnapshotRef.current);
      if (dbVersion.current !== null && dbVersion.current !== undefined) {
        localStorage.setItem(`ecg_db_version:${identity}`, String(dbVersion.current));
      }
    } catch (e) {}
  };

  const getAuthToken = () => sessionStorage.getItem('ecg_session_token');

  const handleUnauthorized = () => {
    sessionStorage.removeItem('ecg_session_token');
    sessionStorage.removeItem('ecg_active_session');
    setCurrentUser(null);
    showToast(language === 'id' ? 'Sesi Anda telah berakhir, silakan masuk kembali.' : 'Your session has expired, please log in again.', 'warning');
  };

  // FUNGSI BARU: Refresh paksa sebelum edit penting (Solusi Anti-Collision)
  const refreshBeforeEdit = async () => {
    const token = getAuthToken();
    if (!token) return handleUnauthorized();

    setSyncStatus('syncing');
    showToast(language === 'id' ? 'Menyinkronkan data terbaru dari server...' : 'Syncing latest data from server...', 'success');
    try {
      const clientVersion = dbVersion.current ?? '';
      const res = await fetch(`${APPSCRIPT_URL}?token=${encodeURIComponent(token)}&clientVersion=${encodeURIComponent(clientVersion)}`);
      const data = await res.json();

      if (data.status === 'unauthorized') return handleUnauthorized();
      if (data.status === 'error') throw new Error(data.message || 'Server sync error');

      if (data.message === 'no_change') {
        setIsCloudConnected(true);
        setSyncStatus('saved');
        showToast(language === 'id' ? 'Data sudah terbaru.' : 'Data is already up to date.', 'success');
        return;
      }

      if (data && data._dbVersion !== undefined && data._dbVersion !== null) {
        dbVersion.current = data._dbVersion;
      }

      const cloudDb = data.payload || data.state_data || data;
      if (cloudDb && Array.isArray(cloudDb.users)) {
        const freshNormalized = normalizeData(cloudDb);
        const localBeforePull = latestDbRef.current;
        const pending = lastSyncedSnapshotRef.current
          ? getPendingSyncOptionsFromBaseline(localBeforePull, lastSyncedSnapshotRef.current)
          : getPendingChangesRelativeToCloud(localBeforePull, freshNormalized);
        const merged = mergeCloudData(localBeforePull, freshNormalized, pending);

        lastSyncedSnapshotRef.current = JSON.stringify(getSyncSnapshot(freshNormalized));
        persistSyncBaseline();
        isDbDirty.current = pending.hasChanges;
        prevEntitiesRef.current = null;
        skipCloudSave.current = !pending.hasChanges;
        setDbState(merged);
        try { localStorage.setItem('ecg_db', JSON.stringify(merged)); } catch (e) {}
        setLogs({
          auditLogs: Array.isArray(cloudDb.auditLogs) ? cloudDb.auditLogs : [],
          debugLogs: Array.isArray(cloudDb.debugLogs) ? cloudDb.debugLogs : []
        });
        setIsCloudConnected(true);
        setSyncStatus(pending.hasChanges ? 'saving' : 'saved');
        showToast(
          pending.hasChanges
            ? (language === 'id' ? 'Data terbaru digabung. Perubahan lokal sedang dikirim...' : 'Latest data merged. Local changes are being uploaded...')
            : (language === 'id' ? 'Data berhasil disinkronkan!' : 'Data synced successfully!'),
          pending.hasChanges ? 'warning' : 'success'
        );
      }
    } catch (error) {
      console.error('Manual refresh failed', error);
      setIsCloudConnected(false);
      setSyncStatus('error');
      showToast(language === 'id' ? 'Gagal menyinkronkan data' : 'Failed to sync data', 'error');
    }
  };

  // FUNGSI BARU: Pelapor Log Langsung ke Sheet
  const sendLogAction = (actionName, details, isError = false) => {
     const activeSession = sessionStorage.getItem('ecg_active_session');
     // BUG FIX #8: Wrap JSON.parse with try-catch — corrupted sessionStorage would crash logging
     let userName = 'SYSTEM/GUEST';
     try { if (activeSession) userName = JSON.parse(activeSession).name || 'SYSTEM/GUEST'; } catch(e) {}
     const token = getAuthToken();

     if (token) {
       // 1. Tembak langsung ke Cloud AppScript
       fetch(APPSCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'log_only', token, user: userName, logAction: actionName, logDetails: details, isError })
       })
       .then(res => res.json())
       .then(data => { if (data.status === 'unauthorized') handleUnauthorized(); })
       .catch(() => {});
     }

     // 2. Show secara instan di UI System Logs Monitor
     const now = new Date();
     const timestamp = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;
     const newLog = { Timestamp: timestamp, User: userName, Action: actionName, Status: isError ? 'FAILED' : 'SUCCESS', Details: details, 'Error Details': isError ? details : undefined };

     setLogs(prev => {
        if (isError) return { ...prev, debugLogs: [newLog, ...(prev.debugLogs || [])].slice(0, 2000) };
        return { ...prev, auditLogs: [newLog, ...(prev.auditLogs || [])].slice(0, 2000) };
     });
  };

  // Diangkat ke atas agar bisa digunakan di mana saja termasuk deteksi tombol Back
  const showToast = (msg, type = 'success') => {
    let finalMsg = msg;
    
    // Sistem Interceptor: Terjemahkan Toast Notifikasi otomatis jika bahasa ID aktif
    if (language === 'id') {
      const toastDict = {
         'You have successfully logged out': 'Anda telah berhasil keluar',
         'Student saved': 'Data siswa tersimpan',
         'Tutor saved': 'Data tutor tersimpan',
         'User Saved': 'Pengguna tersimpan',
         'Attendance Saved': 'Kehadiran tersimpan',
         'Checked in successfully': 'Berhasil absen (Check-in)',
         'Already checked in today': 'Sudah absen hari ini',
         'Updated': 'Berhasil diperbarui',
         'Moved to Recycle Bin': 'Dipindahkan ke Tempat Sampah',
         'Journal saved': 'Jurnal tersimpan',
         'Journal updated': 'Jurnal diperbarui',
         'Payment recorded': 'Pembayaran dicatat',
         'Event created': 'Jadwal dibuat',
         'Event updated': 'Jadwal diperbarui',
         'Published': 'Berhasil dipublikasikan',
         'Item Restored': 'Data dipulihkan',
         'Permanently Deleted': 'Dihapus permanen',
         'Password updated successfully! Welcome to Academic Suite.': 'Sandi berhasil diperbarui! Selamat datang.',
         'Your password has been successfully updated.': 'Sandi Anda berhasil diperbarui.',
         'Cloud connection restored.': 'Koneksi Cloud dipulihkan.',
         'Cloud connection lost. Running in offline mode.': 'Koneksi Cloud terputus. Mode luring aktif.',
         'Material posted successfully': 'Materi berhasil diunggah',
         'Material updated successfully': 'Materi berhasil diperbarui',
         'Submission sent successfully!': 'Tugas berhasil dikirim!',
         'Comment generated automatically': 'Komentar dibuat secara otomatis',
         'Teacher comment saved to profile': 'Komentar guru disimpan ke profil',
         'Press back again to exit': 'Tekan kembali lagi untuk keluar',
         'Please provide either a link or an open-ended question': 'Mohon berikan tautan atau instruksi tugas',
         'You can only manage materials for your assigned session.': 'Anda hanya dapat mengelola materi untuk sesi Anda.',
         'Please write your comment or submission first': 'Silakan tulis komentar atau tugas Anda terlebih dahulu',
         'Schedule ID is required': 'ID Jadwal are required',
         'Enter valid amount': 'Masukkan nominal yang valid',
         'Tutor and Base Salary required.': 'Tutor dan Gaji Pokok are required.',
         'Payroll marked as Paid!': 'Penggajian ditandai Lunas!',
         'Draft Saved!': 'Draf Tersimpan!',
         'Please select a tutor to link': 'Pilih tutor untuk dihubungkan',
         'Please select a student to link': 'Pilih siswa untuk dihubungkan',
         'User deleted successfully': 'Pengguna berhasil dihapus',
         'Your current password is incorrect.': 'Sandi saat ini salah.',
         'New password and confirmation do not match.': 'Sandi baru dan konfirmasi tidak cocok.',
         'Password must be at least 6 characters long.': 'Sandi minimal harus 6 karakter.',
         'Cannot delete the last remaining Super Admin. Create/keep at least one other Super Admin first.': 'Tidak bisa menghapus Super Admin terakhir. Pastikan minimal ada 1 Super Admin lain sebelum menghapus.',
         'Database was updated by another user! Refresh the page to avoid data loss.': 'Database diperbarui pengguna lain! Muat ulang halaman.',
         'Please create an account first to reset password': 'Buat akun terlebih dahulu untuk mereset sandi',
         'No account exists to be deleted.': 'Tidak ada akun untuk dihapus.',
         'Generating PNG...': 'Membuat PNG...',
         'PNG Downloaded!': 'PNG Berhasil Diunduh!',
         'PNG Generation Failed. Try printing to PDF.': 'Gagal membuat PNG. Coba cetak ke PDF.',
         'Preparing image for sharing...': 'Menyiapkan gambar untuk dibagikan...',
         'Shared successfully!': 'Berhasil dibagikan!',
         'Failed to prepare image for sharing.': 'Gagal menyiapkan gambar untuk dibagikan.',
         'No scores entered to save.': 'Tidak ada nilai untuk disimpan.',
         'Image downloaded! (Direct file sharing not supported on this browser)': 'Gambar diunduh! (Berbagi file langsung tidak didukung di peramban ini)'
      };

      if (msg.startsWith('Welcome back, ')) {
         finalMsg = msg.replace('Welcome back, ', 'Selamat datang kembali, ');
      } else if (msg.startsWith('Assessment saved successfully for')) {
         finalMsg = msg.replace('Assessment saved successfully for', 'Penilaian berhasil disimpan untuk').replace('students.', 'siswa.');
      } else if (msg.startsWith('Password reset for')) {
         finalMsg = msg.replace('Password reset for', 'Reset sandi untuk').replace('was successful.', 'berhasil.');
      } else if (msg.includes('item(s) permanently deleted')) {
         finalMsg = msg.replace('item(s) permanently deleted', 'data dihapus permanen');
      } else if (toastDict[msg]) {
         finalMsg = toastDict[msg];
      }
    }

    // PELACAK OTOMATIS: Jika ada notifikasi berhasil simpan/hapus, catat sebagai LOG!
    if (!msg.includes('Welcome back') && !msg.includes('logged out') && !msg.includes('Cloud connection') && !msg.includes('Press back again')) {
       let actionName = 'USER_ACTION';
       const msgLower = msg.toLowerCase();
       
       if (msgLower.includes('delete') || msgLower.includes('hapus') || msgLower.includes('recycle bin') || msgLower.includes('removed')) {
          actionName = 'DELETE_DATA';
       } else if (msgLower.includes('save') || msgLower.includes('simpan') || msgLower.includes('record') || msgLower.includes('publish') || msgLower.includes('update') || msgLower.includes('perbarui')) {
          actionName = 'SAVE_DATA';
       }

       if (type === 'error') {
          sendLogAction('SYSTEM_ERROR', finalMsg, true);
       } else if (type === 'success' || actionName === 'DELETE_DATA') {
          sendLogAction(actionName, finalMsg, false);
       }
    }

    // Bersihkan timer lama jika ada notifikasi baru yang masuk
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    
    // PERBAIKAN: Tambahkan ID unik (Date.now()) agar React memaksa memuat ulang animasi UI-nya
    setToast({ id: Date.now(), msg: finalMsg, type });
    
    // Mulai timer baru
    toastTimeoutRef.current = setTimeout(() => {
      setToast(null);
      toastTimeoutRef.current = null;
    }, 3000);
  };

  // =====================================================================
  // PERBAIKAN 1 & 2: MANAJEMEN SESI (REFRESH) & TOMBOL BACK DI MOBILE
  // =====================================================================
  
  // B. Sinkronisasi Tab Aktif dengan URL Hash (Untuk tombol Back HP)
  const activeTabRef = useRef(activeTab);
  useEffect(() => {
    activeTabRef.current = activeTab;
  }, [activeTab]);

  useEffect(() => {
    const rawHash = window.location.hash.replace('#', '');
    const hash = rawHash.split('?')[0]; // Bersihkan URL dari parameter waktu acak
    if (hash && hash !== activeTab) {
      setActiveTab(hash);
    }
  }, []);

  useEffect(() => {
    const currentHashBase = window.location.hash.replace('#', '').split('?')[0];
    if (currentHashBase !== activeTab) {
      window.history.pushState(null, '', `#${activeTab}`);
    }
  }, [activeTab]);

  useEffect(() => {
    const handlePopState = () => {
      if (isExiting.current) return;

      // LOGIKA BARU: Jika user menekan tombol Back SAAT SEDANG berada di Dashboard
      if (activeTabRef.current === 'dashboard') {
        if (exitToastTimeout.current) {
          // KETUKAN KEDUA: Logout dan kembali ke Login (keluar aplikasi)
          clearTimeout(exitToastTimeout.current);
          exitToastTimeout.current = null;
          
          sessionStorage.removeItem('ecg_active_session');
          localStorage.removeItem('ecg_remembered_user');
          setCurrentUser(null);
          setActiveTab('dashboard');
          window.location.hash = '';
          
          if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
          setToast(null);
          return;
        } else {
          // KETUKAN PERTAMA: Tahan di dashboard dan munculkan peringatan
          // (PushState mengembalikan posisi URL ke dashboard untuk membatalkan back default)
          window.history.pushState(null, '', '#dashboard');
          
          if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
          setToast({ id: Date.now(), msg: 'Press back again to exit', type: 'warning' });
          toastTimeoutRef.current = setTimeout(() => {
            setToast(null);
            toastTimeoutRef.current = null;
          }, 3000);
          
          exitToastTimeout.current = setTimeout(() => {
            exitToastTimeout.current = null;
          }, 2000); // Reset peringatan setelah 2 detik
          return;
        }
      }

      // Jika user BUKAN di dashboard (misal di Students), biarkan back berjalan kembali ke menu/dashboard sebelumnya
      const rawHash = window.location.hash.replace('#', '');
      const hash = rawHash.split('?')[0];

      if (hash && hash !== 'dashboard') {
        setActiveTab(hash);
        if (exitToastTimeout.current) {
           clearTimeout(exitToastTimeout.current);
           exitToastTimeout.current = null;
        }
      } else {
        setActiveTab('dashboard');
      }
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  // PATCH 3: CLEAR TOAST TIMEOUT SAAT PINDAH MENU
  useEffect(() => {
    if (activeTab !== 'dashboard' && exitToastTimeout.current) {
      clearTimeout(exitToastTimeout.current);
      exitToastTimeout.current = null;
    }
  }, [activeTab]);
  // =====================================================================

  // Set browser tab title
  useEffect(() => {
    document.title = "ECG Academic Suite";
  }, []);

  useEffect(() => {
    const loadData = async () => {
      // Pulihkan baseline sync dan versi server per-account. Baseline ini diperlukan
      // agar deletion offline dapat dibedakan dari cache baru/parsial yang belum pernah tersinkron.
      try {
        let identity = 'anonymous';
        const active = sessionStorage.getItem('ecg_active_session');
        if (active) {
          const parsed = JSON.parse(active);
          const rawIdentity = parsed?.id || parsed?.username || parsed?.name;
          if (rawIdentity) identity = String(rawIdentity).trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_');
        }
        const persistedSnapshot = localStorage.getItem(`ecg_sync_snapshot:${identity}`);
        if (persistedSnapshot) lastSyncedSnapshotRef.current = persistedSnapshot;
        const persistedVersion = Number(localStorage.getItem(`ecg_db_version:${identity}`));
        if (Number.isFinite(persistedVersion) && persistedVersion >= 0) dbVersion.current = persistedVersion;
      } catch (e) {}

      // 1. MUAT DATA LOKAL LEBIH DULU AGAR INSTAN
      const saved = localStorage.getItem('ecg_db');
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (!parsed.users || !Array.isArray(parsed.users)) {
            skipCloudSave.current = true;
            setDbState(generateDummyDatabase());
          } else {
            skipCloudSave.current = true;
            const normLocal = normalizeData(parsed);
            setDbState(normLocal);
            // FIX: Muat logs dari localStorage ke state terpisah
            setLogs({
              auditLogs: Array.isArray(parsed.auditLogs) ? parsed.auditLogs : [],
              debugLogs: Array.isArray(parsed.debugLogs) ? parsed.debugLogs : []
            });
          }
        } catch (e) {
          skipCloudSave.current = true;
          setDbState(generateDummyDatabase());
        }
      } else {
        skipCloudSave.current = true;
        setDbState(generateDummyDatabase());
      }
      
      // 2. LANGSUNG AKTIFKAN APLIKASI & TOMBOL LOGIN (Tanpa Menunggu Cloud)
      setIsDbLoaded(true);

      // 3. AMBIL DATA CLOUD DI LATAR BELAKANG — hanya jika sudah ada token login asli.
      // Backend sekarang WAJIB token valid (hasil security audit): token palsu/pre_auth
      // selalu ditolak 'unauthorized'. Kalau belum login, tidak ada token → skip fetch.
      // Data cloud akan diambil otomatis di finalizeLogin setelah login berhasil.
      const token = getAuthToken();
      if (!token) {
        // Belum ada sesi aktif — data lokal di atas sudah cukup untuk tampilkan UI login.
        return;
      }

      try {
        const startupController = new AbortController();
        const startupTimeout = setTimeout(() => startupController.abort(), 12000);
        const res = await fetch(`${APPSCRIPT_URL}?token=${encodeURIComponent(token)}&clientVersion=${encodeURIComponent(dbVersion.current ?? '')}`, { signal: startupController.signal });
        clearTimeout(startupTimeout);
        if (!res.ok) throw new Error('Failed to load from AppScript');
        
        const data = await res.json();
        
        if (data.status === 'unauthorized') return handleUnauthorized();

        if (data && data.status === 'error') {
           throw new Error(data.message || 'Internal Server Error from AppScript');
        }

        // NEW: Ambil _dbVersion dari root level response (sesuai struktur AppScript)
        if (data && data._dbVersion) {
           dbVersion.current = data._dbVersion;
        }

        const cloudDb = data.payload || data.state_data || data;

        if (data.message === 'no_change') {
          setIsCloudConnected(true);
          setSyncStatus('saved');
          return;
        }

        if (cloudDb && Array.isArray(cloudDb.users)) {
          const normalizedCloudStartup = normalizeData(cloudDb);
          const localBeforePull = latestDbRef.current;
          const pending = lastSyncedSnapshotRef.current
            ? getPendingSyncOptionsFromBaseline(localBeforePull, lastSyncedSnapshotRef.current)
            : getPendingChangesRelativeToCloud(localBeforePull, normalizedCloudStartup);
          const mergedStartup = mergeCloudData(localBeforePull, normalizedCloudStartup, pending);

          // Cloud snapshot menjadi baseline; hanya delta lokal yang benar-benar berbeda
          // dari cloud yang akan dilanjutkan ke proses sync.
          lastSyncedSnapshotRef.current = JSON.stringify(getSyncSnapshot(normalizedCloudStartup));
          persistSyncBaseline();
          isDbDirty.current = pending.hasChanges;
          prevEntitiesRef.current = null;
          skipCloudSave.current = !pending.hasChanges;
          setDbState(mergedStartup);
          try { localStorage.setItem('ecg_db', JSON.stringify(mergedStartup)); } catch (e) {}
          setLogs({
            auditLogs: Array.isArray(cloudDb.auditLogs) ? cloudDb.auditLogs : [],
            debugLogs: Array.isArray(cloudDb.debugLogs) ? cloudDb.debugLogs : []
          });
          setIsCloudConnected(true);
          setSyncStatus(pending.hasChanges ? 'saving' : 'saved');
        } else {
          console.warn('Empty or invalid data format from AppScript', data);
          setIsCloudConnected(false);
        }
      } catch (e) {
        console.warn('AppScript connection failed, using Local Storage', e);
        setIsCloudConnected(false);
      }
    };
    loadData();
  }, []);

  useEffect(() => {
    if (isDbLoaded && db && Array.isArray(db.users)) {
      // Selalu simpan cache lokal terlebih dahulu.
      try { localStorage.setItem('ecg_db', JSON.stringify(db)); } catch (e) {}

      // Pull awal dari local/cloud tidak boleh dianggap mutation user.
      if (skipCloudSave.current) {
        skipCloudSave.current = false;
        prevEntitiesRef.current = JSON.stringify(getSyncSnapshot(db));
        return;
      }

      const currentEntities = JSON.stringify(getSyncSnapshot(db));
      if (prevEntitiesRef.current === currentEntities) return;
      prevEntitiesRef.current = currentEntities;

      isDbDirty.current = true;
      setSyncStatus('saving');

      if (syncDebounceTimer.current) clearTimeout(syncDebounceTimer.current);

      syncDebounceTimer.current = setTimeout(() => {
        const token = getAuthToken();
        if (!token) return handleUnauthorized();

        const latestDb = latestDbRef.current;
        const lastSnap = (() => {
          if (!lastSyncedSnapshotRef.current) return null;
          try { return JSON.parse(lastSyncedSnapshotRef.current); } catch (e) { return null; }
        })();
        const diff = buildSyncDelta(latestDb, lastSnap);

        // Tidak ada delta nyata = tidak perlu menaikkan versi server.
        if (!diff.hasChanges) {
          isDbDirty.current = false;
          activeSyncRequestIdRef.current = null;
          activeSyncFingerprintRef.current = null;
          setSyncStatus('saved');
          return;
        }

        const syncFingerprint = JSON.stringify({
          baseVersion: dbVersion.current,
          deltaPayload: diff.deltaPayload,
          deletions: diff.deletions
        });
        let requestId = activeSyncRequestIdRef.current;
        if (!requestId || activeSyncFingerprintRef.current !== syncFingerprint) {
          requestId = `sync-${Date.now().toString(36)}-${(++syncRequestCounterRef.current).toString(36)}-${Math.random().toString(36).slice(2,10)}`;
          activeSyncRequestIdRef.current = requestId;
          activeSyncFingerprintRef.current = syncFingerprint;
        }
        setSyncStatus('syncing');

        const controller = new AbortController();
        const requestTimeout = setTimeout(() => controller.abort(), 20000);

        fetch(APPSCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          redirect: 'follow',
          signal: controller.signal,
          body: JSON.stringify({
            action: 'sync',
            token,
            requestId,
            baseVersion: dbVersion.current,
            user: currentUser ? currentUser.name : 'SYSTEM',
            payload: diff.deltaPayload,
            deletions: diff.deletions
          })
        })
        .then(res => {
          clearTimeout(requestTimeout);
          if (!res.ok) throw new Error(`Response AppScript gagal (${res.status})`);
          return res.json().catch(() => ({}));
        })
        .then(data => {
          if (data.status === 'unauthorized') {
            activeSyncRequestIdRef.current = null;
            activeSyncFingerprintRef.current = null;
            return handleUnauthorized();
          }

          if (data.status === 'success') {
            if (data.newVersion !== undefined && data.newVersion !== null) {
              dbVersion.current = data.newVersion;
            }

            // Snapshot server = state client yang baru saja dikirim. Dengan strict baseVersion,
            // tidak ada device lain yang dapat menyisipkan write di antara read dan commit ini.
            lastSyncedSnapshotRef.current = JSON.stringify(diff.dbSnapshotAtRequest);
            activeSyncRequestIdRef.current = null;
            activeSyncFingerprintRef.current = null;
            syncBusyAttempt.current = 0;
            setIsCloudConnected(true);
            setSyncStatus('saved');

            const latestAfterResponse = latestDbRef.current;
            const stillChanged = JSON.stringify(getSyncSnapshot(latestAfterResponse)) !== JSON.stringify(diff.dbSnapshotAtRequest);
            if (stillChanged) {
              isDbDirty.current = true;
              prevEntitiesRef.current = null;
            } else {
              isDbDirty.current = false;
            }
            return;
          }

          if (data.status === 'conflict') {
            console.warn('DATABASE CONFLICT — merging server state and retrying only local delta');
            showToast(
              language === 'id'
                ? 'Data diperbarui pengguna lain. Menggabungkan perubahan Anda...'
                : 'Data was updated by another device. Merging your changes...',
              'warning'
            );

            const freshPayload = data.payload;
            const freshVersion = data.newVersion ?? data._dbVersion ?? null;
            if (!freshPayload) throw new Error('Conflict response tidak membawa payload server.');

            const freshNormalized = normalizeData(freshPayload);
            const localPriorityByCollection: Record<string, Set<string>> = {};
            const deletedIdsByCollection: Record<string, Set<string>> = {};
            Object.keys(diff.deltaPayload).forEach(col => {
              const ids = new Set<string>();
              (diff.deltaPayload[col] || []).forEach(item => {
                const id = getRecordId(item);
                if (id) ids.add(id);
              });
              if (ids.size) localPriorityByCollection[col] = ids;
            });
            diff.deletions.forEach(d => {
              if (!deletedIdsByCollection[d.collection]) deletedIdsByCollection[d.collection] = new Set<string>();
              deletedIdsByCollection[d.collection].add(String(d.id));
            });

            const conflictMerged = mergeCloudData(
              latestDbRef.current,
              freshNormalized,
              { localPriorityByCollection, deletedIdsByCollection }
            );

            if (freshVersion !== null && freshVersion !== undefined) dbVersion.current = freshVersion;
            // Server snapshot menjadi baseline baru; local delta akan dihitung ulang dari sini.
            lastSyncedSnapshotRef.current = JSON.stringify(getSyncSnapshot(freshNormalized));
            persistSyncBaseline();
            const retryDiff = buildSyncDelta(conflictMerged, freshNormalized);

            activeSyncRequestIdRef.current = null;
            activeSyncFingerprintRef.current = null;
            syncBusyAttempt.current = 0;
            isDbDirty.current = retryDiff.hasChanges;
            prevEntitiesRef.current = null;
            // Perubahan cloud tidak boleh dikirim ulang jika tidak ada local delta.
            skipCloudSave.current = !retryDiff.hasChanges;
            setDbState(conflictMerged);
            setIsCloudConnected(true);
            setSyncStatus(retryDiff.hasChanges ? 'saving' : 'saved');
            return;
          }

          if (data.status === 'busy') {
            syncBusyAttempt.current += 1;
            const attempt = syncBusyAttempt.current;
            const delay = Math.min(2000 * Math.pow(2, attempt - 1), 16000);
            console.warn(`AppScript busy — retry #${attempt} dalam ${delay/1000}s`);
            setSyncStatus('saving');
            setTimeout(() => {
              if (syncBusyAttempt.current >= attempt) syncBusyAttempt.current = 0;
              isDbDirty.current = true;
              prevEntitiesRef.current = null;
              // Ref saja tidak memicu React render; shallow-copy raw state untuk memicu retry.
              setDbState(prev => ({ ...prev }));
            }, delay);
            return;
          }

          throw new Error(data.message || 'Sync error');
        })
        .catch(e => {
          clearTimeout(requestTimeout);
          console.warn('AppScript Sync failed', e);
          // JANGAN hapus requestId. Bila server sebenarnya sudah commit tetapi respons hilang,
          // retry dengan requestId yang sama akan diambil dari cache idempotency di GAS.
          setIsCloudConnected(false);
          setSyncStatus('error');
          syncBusyAttempt.current = 0;
          isDbDirty.current = true;
          prevEntitiesRef.current = null;
        });
      }, 2000);
    }
  }, [db, isDbLoaded]);

  // Effect khusus untuk memunculkan notifikasi perubahan status cloud
  useEffect(() => {
    if (prevCloudState.current !== isCloudConnected) {
      if (isCloudConnected) {
        showToast('Cloud connection restored.', 'success');
      } else {
        showToast('Cloud connection lost. Running in offline mode.', 'warning');
      }
      prevCloudState.current = isCloudConnected;
    }
  }, [isCloudConnected]);

  // FIX SYNC ERROR: Auto-retry saat koneksi terputus (interval polling 30 detik)
  // Diperlukan karena ketika sync gagal, tidak ada trigger otomatis untuk coba ulang
  // kecuali user mengubah data. Ini memastikan reconect terjadi tanpa interaksi user.
  const offlineRetryTimer = useRef(null);
  useEffect(() => {
    if (!isCloudConnected) {
      // Mulai polling setiap 30 detik selama offline
      offlineRetryTimer.current = setInterval(() => {
        const token = getAuthToken();
        if (!token) return; // Belum login, skip
        // Paksa re-sync dengan menandai dirty + null-kan snapshot pembanding
        if (!isDbDirty.current) {
          isDbDirty.current = true;
          prevEntitiesRef.current = null;
          // Trigger useEffect sync dengan shallow copy db
          setDbState(prev => ({ ...prev }));
        }
      }, 30000);
    } else {
      // Koneksi pulih, hentikan polling
      if (offlineRetryTimer.current) {
        clearInterval(offlineRetryTimer.current);
        offlineRetryTimer.current = null;
      }
    }
    return () => {
      if (offlineRetryTimer.current) clearInterval(offlineRetryTimer.current);
    };
  }, [isCloudConnected]);

  // FIX SYNC ERROR: Dengarkan event browser 'online' untuk trigger sync segera
  // saat koneksi internet pulih tanpa menunggu interval 30 detik
  useEffect(() => {
    const handleOnline = () => {
      if (!isCloudConnected) {
        const token = getAuthToken();
        if (!token) return;
        isDbDirty.current = true;
        prevEntitiesRef.current = null;
        setDbState(prev => ({ ...prev }));
      }
    };
    window.addEventListener('online', handleOnline);
    return () => window.removeEventListener('online', handleOnline);
  }, [isCloudConnected]);

  const requestConfirm = (title, message, onConfirm) => {
    setConfirmDialog({
      title,
      message,
      onConfirm: () => {
        onConfirm();
        setConfirmDialog(null);
      },
    });
  };

  const softDelete = (collection, id, itemName) => {
    requestConfirm(
      'Confirm Deletion',
      `Are you sure you want to delete ${itemName || 'this record'}? It will be moved to the Recycle Bin.`,
      () => {
        const item = db[collection].find((x) => x.id === id);
        // FIX BUG #1: binItem HARUS punya field 'id' agar mergeByIds tidak memfilternya
        // (isInvalid() mengembalikan true jika !item.id — binItem lama hanya pakai 'binId')
        const binId = `BIN-${Date.now()}`;
        const binItem = {
          id: binId,          // ← primary key untuk mergeByIds
          binId: binId,       // dipertahankan untuk kompatibilitas UI (restore/delete Recycle Bin)
          originalCollection: collection,
          deletedAt: getSyncTimestamp(),
          data: item,
        };
        setDb((prev) => ({
          ...prev,
          [collection]: prev[collection].filter((x) => x.id !== id),
          recycleBin: [...(prev.recycleBin || []), binItem],
        }));
        showToast('Moved to Recycle Bin', 'warning');
      }
    );
  };

  const generateId = (prefix, collection) => {
    const generateRandomPart = () => {
      // 4 digit angka acak (1000-9999)
      const digits = Math.floor(1000 + Math.random() * 9000).toString();
      
      // 4 huruf kapital acak (A-Z)
      let letters = '';
      const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
      for (let i = 0; i < 4; i++) {
        letters += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
      }
      
      return `${digits}-${letters}`;
    };

    let newId;
    let isUnique = false;
    // FIX #16: Batasi iterasi maksimal 100x untuk mencegah infinite loop jika
    // prefix space terlalu penuh. Jika melebihi batas, tambahkan timestamp sebagai suffix.
    const MAX_ATTEMPTS = 100;
    let attempt = 0;

    // Loop pencegah collision: Memastikan ID benar-benar unik di seluruh database
    while (!isUnique && attempt < MAX_ATTEMPTS) {
      attempt++;
      newId = `${prefix}-${generateRandomPart()}`;
      
      // Periksa apakah ID sudah terpakai di koleksi data aktif
      const existsInDb = (db[collection] || []).some(item => item.id === newId);
      
      // Periksa apakah ID sudah terpakai di Recycle Bin (mencegah duplikasi data yang sudah dihapus)
      const existsInBin = (db.recycleBin || []).some(
        binItem => binItem.originalCollection === collection && binItem.data?.id === newId
      );

      if (!existsInDb && !existsInBin) {
        isUnique = true;
      }
    }

    // Fallback jika semua percobaan gagal (sangat jarang): tambahkan timestamp unik
    if (!isUnique) {
      newId = `${prefix}-${generateRandomPart()}-${Date.now().toString(36).toUpperCase()}`;
    }

    return newId;
  };

  // LOGIN HANDLER: Login standar menggunakan database lokal (Hardcoded Admin/Password)
  const handleLogin = async (username, password, rememberMe) => {
    const cleanUser = String(username).trim().toLowerCase();
    const cleanPass = String(password).trim();

    const finalizeLogin = (userObj, sessionToken) => {
       sessionStorage.setItem('ecg_session_token', sessionToken);
       sessionStorage.setItem('ecg_active_session', JSON.stringify(userObj));
       if (rememberMe) localStorage.setItem('ecg_remembered_user', JSON.stringify({ username: cleanUser }));
       else localStorage.removeItem('ecg_remembered_user');
       setCurrentUser(userObj);
       setSavedAccounts(prev => {
          const existing = prev.filter(acc => String(acc.username || '').toLowerCase() !== cleanUser);
          const newAccounts = [...existing, { username: cleanUser, name: userObj.name, role: userObj.role, lastToken: sessionToken }];
          localStorage.setItem('ecg_saved_accounts', JSON.stringify(newAccounts));
          return newAccounts;
       });
       sendLogAction('LOGIN', `${userObj.role} logged in`);
       showToast(language === 'id' ? `Selamat datang kembali, ${userObj.name}` : `Welcome back, ${userObj.name}`);
       return { success: true };
    };

    const findUserInDb = (dbToSearch) => {
      let found = (dbToSearch.users || []).find(u =>
        String(u.username || '').trim().toLowerCase() === cleanUser &&
        String(u.password || '').trim() === cleanPass
      );
      if (!found) {
        const tutor = (dbToSearch.tutors || []).find(t =>
          String(t.username || '').trim().toLowerCase() === cleanUser &&
          String(t.password || '').trim() === cleanPass
        );
        if (tutor) found = { ...tutor, role: 'tutor' };
      }
      return found;
    };

    // ─── FAST PATH: try local cache immediately (< 5ms) ────────────────────
    const localUser = findUserInDb(db);
    if (localUser) {
      if (localUser.active !== 'Active' && localUser.status !== 'Active') {
        return { success: false, error: language === 'id' ? 'Akun Anda tidak aktif.' : 'Your account is inactive.' };
      }
      const sessionUser = { ...localUser };
      if (String(sessionUser.role).toLowerCase().includes('super')) {
        sessionUser.role = 'admin';
        sessionUser.isSuperAdmin = true;
      }

      // Log in immediately WITHOUT storing a token.
      // No token = startup fetch skips GAS = no false "unauthorized" = no offline flag.
      // Real token arrives from background verify below.
      sessionStorage.setItem('ecg_active_session', JSON.stringify(sessionUser));
      sessionStorage.removeItem('ecg_session_token');
      if (rememberMe) localStorage.setItem('ecg_remembered_user', JSON.stringify({ username: cleanUser }));
      else localStorage.removeItem('ecg_remembered_user');
      setCurrentUser(sessionUser);
      // Optimistic: assume online until background verify proves otherwise
      setIsCloudConnected(true);
      setSyncStatus('syncing');
      setSavedAccounts(prev => {
        const existing = prev.filter(acc => String(acc.username || '').toLowerCase() !== cleanUser);
        return [...existing, { username: cleanUser, name: sessionUser.name, role: sessionUser.role, lastToken: '' }];
      });
      sendLogAction('LOGIN', `${sessionUser.role} logged in (fast-path)`);
      showToast(language === 'id' ? `Selamat datang, ${sessionUser.name}` : `Welcome, ${sessionUser.name}`);

      // ─── BACKGROUND: get real token + sync data ────────────────────────────
      const controller = new AbortController();
      const bgTimeout = setTimeout(() => {
        controller.abort();
        setIsCloudConnected(false);
        setSyncStatus('error');
      }, 12000);

      fetch(APPSCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        redirect: 'follow',
        signal: controller.signal,
        body: JSON.stringify({ action: 'login', username: cleanUser, password: cleanPass })
      })
        .then(res => { clearTimeout(bgTimeout); return res.json().catch(() => null); })
        .then(result => {
          if (result?.status === 'success' && result.token) {
            sessionStorage.setItem('ecg_session_token', result.token);
            setSavedAccounts(prev => {
              const existing = prev.filter(acc => String(acc.username || '').toLowerCase() !== cleanUser);
              return [...existing, { username: cleanUser, name: sessionUser.name, role: sessionUser.role, lastToken: result.token }];
            });
            setIsCloudConnected(true);
            const syncController = new AbortController();
            const syncTimeout = setTimeout(() => syncController.abort(), 15000);
            fetch(`${APPSCRIPT_URL}?token=${encodeURIComponent(result.token)}&clientVersion=${encodeURIComponent(dbVersion.current ?? '')}`, { signal: syncController.signal })
              .then(r => { clearTimeout(syncTimeout); return r.json(); })
              .then(data => {
                const cloudDb = data.payload || data.state_data || data;
                if (data.message === 'no_change') {
                  setIsCloudConnected(true);
                  setSyncStatus('saved');
                  return;
                }
                if (cloudDb && Array.isArray(cloudDb.users)) {
                  if (data._dbVersion !== undefined && data._dbVersion !== null) dbVersion.current = data._dbVersion;
                  const normalizedCloud = normalizeData(cloudDb);
                  const localBeforePull = latestDbRef.current;
                  const pending = lastSyncedSnapshotRef.current
                    ? getPendingSyncOptionsFromBaseline(localBeforePull, lastSyncedSnapshotRef.current)
                    : getPendingChangesRelativeToCloud(localBeforePull, normalizedCloud);
                  const merged = mergeCloudData(localBeforePull, normalizedCloud, pending);
                  lastSyncedSnapshotRef.current = JSON.stringify(getSyncSnapshot(normalizedCloud));
                  persistSyncBaseline();
                  isDbDirty.current = pending.hasChanges;
                  prevEntitiesRef.current = null;
                  skipCloudSave.current = !pending.hasChanges;
                  setDbState(merged);
                  try { localStorage.setItem('ecg_db', JSON.stringify(merged)); } catch (e) {}
                  setLogs({ auditLogs: Array.isArray(cloudDb.auditLogs) ? cloudDb.auditLogs : [], debugLogs: Array.isArray(cloudDb.debugLogs) ? cloudDb.debugLogs : [] });
                  setSyncStatus(pending.hasChanges ? 'saving' : 'saved');
                }
              })
              .catch(() => { clearTimeout(syncTimeout); setSyncStatus('error'); });
          } else if (result?.status === 'error') {
            // BUGFIX: Jangan logout user saat background verify gagal.
            // User sudah berhasil login via data lokal (fast-path).
            // Error server hanya berarti sinkronisasi cloud tidak tersedia,
            // bukan alasan untuk membatalkan sesi yang sudah aktif.
            console.warn('Background token verification failed — staying logged in with local session.');
            setIsCloudConnected(false);
            setSyncStatus('error');
          } else {
            setIsCloudConnected(false);
            setSyncStatus('error');
          }
        })
        .catch(() => {
          clearTimeout(bgTimeout);
          setIsCloudConnected(false);
          setSyncStatus('error');
        });

      return { success: true };
    }

    // ─── SLOW PATH: no local cache — must wait for server ───────────────────
    // (first-time login on a new device, or account just created by admin)
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 8000);
      const res = await fetch(APPSCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        redirect: 'follow',
        signal: controller.signal,
        body: JSON.stringify({ action: 'login', username: cleanUser, password: cleanPass })
      });
      clearTimeout(timeout);
      const result = await res.json().catch(() => null);

      if (result?.status === 'success' && result.token && result.user) {
        const sessionUser = { ...result.user };
        if (String(sessionUser.role).toLowerCase().includes('super')) {
          sessionUser.role = 'admin';
          sessionUser.isSuperAdmin = true;
        }
        setIsCloudConnected(true);
        const loginResult = finalizeLogin(sessionUser, result.token);
        // Background sync after first-time login
        setSyncStatus('syncing');
        const syncController = new AbortController();
        const syncTimeout = setTimeout(() => syncController.abort(), 15000);
        fetch(`${APPSCRIPT_URL}?token=${encodeURIComponent(result.token)}&clientVersion=${encodeURIComponent(dbVersion.current ?? '')}`, { signal: syncController.signal })
          .then(r => { clearTimeout(syncTimeout); return r.json(); })
          .then(data => {
            const cloudDb = data.payload || data.state_data || data;
            if (data.message === 'no_change') {
              setIsCloudConnected(true);
              setSyncStatus('saved');
            } else if (cloudDb && Array.isArray(cloudDb.users)) {
              if (data._dbVersion !== undefined && data._dbVersion !== null) dbVersion.current = data._dbVersion;
              const normalizedCloud = normalizeData(cloudDb);
              const localBeforePull = latestDbRef.current;
              const pending = lastSyncedSnapshotRef.current
              ? getPendingSyncOptionsFromBaseline(localBeforePull, lastSyncedSnapshotRef.current)
              : getPendingChangesRelativeToCloud(localBeforePull, normalizedCloud);
              const merged = mergeCloudData(localBeforePull, normalizedCloud, pending);
              lastSyncedSnapshotRef.current = JSON.stringify(getSyncSnapshot(normalizedCloud));
              persistSyncBaseline();
              isDbDirty.current = pending.hasChanges;
              prevEntitiesRef.current = null;
              skipCloudSave.current = !pending.hasChanges;
              setDbState(merged);
              try { localStorage.setItem('ecg_db', JSON.stringify(merged)); } catch (e) {}
              setLogs({ auditLogs: Array.isArray(cloudDb.auditLogs) ? cloudDb.auditLogs : [], debugLogs: Array.isArray(cloudDb.debugLogs) ? cloudDb.debugLogs : [] });
              setSyncStatus(pending.hasChanges ? 'saving' : 'saved');
            }
          })
          .catch(() => { clearTimeout(syncTimeout); setSyncStatus('error'); });
        return loginResult;
      }

      if (result?.status === 'error') {
        return { success: false, error: result.message || (language === 'id' ? 'Nama pengguna atau kata sandi salah.' : 'Invalid username or password.') };
      }
    } catch (e) {
      console.warn('Server login unreachable', e);
    }

    return { success: false, error: language === 'id' ? 'Tidak dapat terhubung ke server. Periksa koneksi internet Anda.' : 'Cannot connect to server. Please check your internet connection.' };
  };

  const handleLogout = () => {
    setLogoutConfirm(true);
  };

  const confirmLogout = () => {
    sendLogAction('LOGOUT', 'User logged out manually'); // PENCATATAN LOG
    // PATCH 2: Hapus semua sesi dari storage saat logout
    sessionStorage.removeItem('ecg_session_token');
    sessionStorage.removeItem('ecg_active_session');
    localStorage.removeItem('ecg_remembered_user');
    
    setCurrentUser(null);
    setActiveTab('dashboard');
    window.location.hash = ''; // Bersihkan URL Hash
    setLogoutConfirm(false);
    sonnerToast.success('You have successfully logged out');
  };

  const switchAccount = async (acc) => {
    // SECURITY FIX: Password tidak lagi disimpan di localStorage.
    // Switch account langsung arahkan ke login screen dengan username pre-filled.
    sessionStorage.removeItem('ecg_session_token');
    sessionStorage.removeItem('ecg_active_session');
    localStorage.setItem('ecg_remembered_user', JSON.stringify({ username: acc.username }));
    setCurrentUser(null);
    setShowAccountMenu(false);
    setActiveTab('dashboard');
    showToast(language === 'id' ? `Silakan masukkan password untuk ${acc.name}` : `Please enter password for ${acc.name}`, 'info');
  };

  const removeSavedAccount = (e, username) => {
    e.stopPropagation();
    setSavedAccounts(prev => {
       // BUG FIX #6: Use case-insensitive comparison to handle pre-migration accounts with mixed case
       const newAcc = prev.filter(a => String(a.username || '').toLowerCase() !== String(username || '').toLowerCase());
       localStorage.setItem('ecg_saved_accounts', JSON.stringify(newAcc));
       return newAcc;
    });
  };

  const downloadPNG = async (elementId, filename) => {
    showToast('Generating PNG...', 'success');
    try {
      if (!window.html2canvas) {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        document.body.appendChild(script);
        // BUG FIX #7: Add onerror handler to prevent infinite hang when CDN is unreachable
        await new Promise((resolve, reject) => {
          script.onload = resolve;
          script.onerror = () => reject(new Error('Failed to load html2canvas from CDN'));
        });
      }
      const element = document.getElementById(elementId);
      const canvas = await window.html2canvas(element, { scale: 2, backgroundColor: '#ffffff', useCORS: true });
      const imgData = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.download = `${filename}.png`;
      link.href = imgData;
      link.click();
      showToast('PNG Downloaded!');
    } catch (err) {
      console.error(err);
      showToast('PNG Generation Failed. Try printing to PDF.', 'error');
    }
  };

  const handleShareImage = async (elementId, filename, text) => {
    showToast('Preparing image for sharing...', 'success');
    try {
      if (!window.html2canvas) {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        document.body.appendChild(script);
        // BUG FIX #7: Add onerror handler to prevent infinite hang when CDN is unreachable
        await new Promise((resolve, reject) => {
          script.onload = resolve;
          script.onerror = () => reject(new Error('Failed to load html2canvas from CDN'));
        });
      }
      const element = document.getElementById(elementId);
      const canvas = await window.html2canvas(element, { scale: 2, backgroundColor: '#ffffff', useCORS: true });
      
      canvas.toBlob(async (blob) => {
        const file = new File([blob], `${filename}.png`, { type: blob.type });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          try {
            await navigator.share({
              title: filename,
              text: text,
              files: [file]
            });
            showToast('Shared successfully!');
          } catch (err) {
            // Share dibatalkan pengguna atau gagal secara diam-diam (fallback ke download di bawah bila perlu)
          }
        } else {
           const imgData = canvas.toDataURL('image/png');
           const link = document.createElement('a');
           link.download = `${filename}.png`;
           link.href = imgData;
           link.click();
           showToast('Image downloaded! (Direct file sharing not supported on this browser)', 'warning');
        }
      });
    } catch (err) {
       console.error(err);
           showToast('Failed to prepare image for sharing.', 'error');
    }
  };

  // CALCULATE BADGES FOR TUTOR & STUDENT NOTIFICATIONS
  const badgeCounts = useMemo(() => {
    let counts = { 
      tutor_attendance: 0, journals: 0, assessments: 0, student_attendance: 0, materials: 0,
      my_payments: 0, my_materials: 0, calendar: 0, announcements: 0, speaking_challenge: 0
    };
    
    const today = getTodayDateLocal();
    const dObj = new Date();
    const currentMonth = String(dObj.getMonth() + 1);
    const currentYear = String(dObj.getFullYear());

    if (currentUser?.role === 'tutor') {
      const isMyClass = (tutorString, myName) => tutorString && tutorString.split(' & ').includes(myName);
      // Fix #11: kecualikan event dibatalkan dari hitungan badge notifikasi
      const CANCELLED_TYPES_BADGE = ['Cancelled', 'Holiday', 'Off Day', 'Libur', 'Dibatalkan'];
      const mySchedulesToday = db.calendar.filter(
        c => c.date === today && isMyClass(c.tutor, currentUser.name) && !CANCELLED_TYPES_BADGE.includes(c.type)
      );
      
      // 1. Tutor Check-In (Jika ada jadwal kelas hari ini tapi belum check-in)
      // Fix #4: fallback ke name untuk data legacy yang belum punya tutorId
      const hasCheckedIn = db.tutorAttendance.some(
        a => (a.tutorId === currentUser.id || (!a.tutorId && a.name === currentUser.name)) && a.date === today && a.status === 'Present'
      );
      counts.tutor_attendance = (!hasCheckedIn && mySchedulesToday.length > 0) ? 1 : 0;
      
      // 2. Learning Journals (Kurangi jumlah jadwal dengan jumlah jurnal hari ini)
      const journalsToday = db.journals.filter(j => j.tutorName === currentUser.name && j.date === today).length;
      counts.journals = Math.max(0, mySchedulesToday.length - journalsToday);
      
      // 3. Student Attendance (Jadwal hari ini yang belum disubmit absensinya)
      const markedScheduleIdsToday = Array.from(new Set(db.studentAttendance.filter(a => a.date === today).map(a => a.scheduleId)));
      counts.student_attendance = mySchedulesToday.filter(c => !markedScheduleIdsToday.includes(c.id)).length;
      
      // 4. Monthly Assessments (Jumlah murid aktif - jumlah form nilai yang disubmit bulan ini)
      const activeStudents = db.students.filter(s => s.status === 'Active');
      const mySessions = parseSessions(currentUser.teachingSession);
      const myStudents = activeStudents.filter(s => mySessions.includes(getStudentSession(s))).length;
      // FIX #13: Deduplikasi studentId agar satu siswa dengan 2 record assessment
      // tidak menggelembungkan assessmentsDone melebihi myStudents → badge negatif.
      const assessmentsDone = new Set(
        db.assessments
          .filter(a => Number(a.month) === Number(currentMonth) && String(a.year) === String(currentYear) && mySessions.includes(a.sessionGroup))
          .map(a => a.studentId)
      ).size;
      counts.assessments = Math.max(0, myStudents - assessmentsDone);
      
      // 5. Materials & Tasks (Menghitung submission / komentar siswa yang belum ditandai checked/read)
      let pendingMaterialsCount = 0;
      (db.materials || []).filter(m => parseSessions(currentUser.teachingSession).includes(m.sessionGroup)).forEach(mat => {
         pendingMaterialsCount += (mat.submissions || []).filter(sub => !sub.checked).length;
      });
      counts.materials = pendingMaterialsCount;

    } else if (currentUser?.role === 'student') {
      const studentRec = db.students.find(s => s.id === currentUser.studentId);
      const mySession = studentRec ? getStudentSession(studentRec) : '';

      // 1. My Payments: Munculkan 1 sebagai PENGINGAT TAGIHAN jika SPP bulan ini belum dibayar penuh
      const plan = studentRec?.paymentPlan || 'Monthly';
      const monthPrefix = `${currentYear}-${currentMonth.padStart(2, '0')}`;
      let target = 0;
      if (plan === 'Free') {
          target = 0; // FREE: tidak pernah ada tagihan
      } else if (plan === 'Monthly') {
          target = db.calendar.filter(c => c.date.startsWith(monthPrefix) && (c.sessionGroup || c.name) === mySession).length * 25000;
      } else {
          target = db.studentAttendance.filter(a => a.studentId === currentUser.studentId && a.date.startsWith(monthPrefix) && a.status === 'Present').length * 25000;
      }
      const totalPaid = db.payments.filter(p => p.studentId === currentUser.studentId && Number(p.month) === Number(currentMonth) && String(p.year) === String(currentYear) && p.status === 'Paid').reduce((sum, p) => sum + Number(p.amount), 0);
      
      counts.my_payments = (totalPaid < target && target > 0) ? 1 : 0;

      // 2. My Materials: Jumlah SEMUA tugas/materi yang belum dikerjakan (submit) oleh siswa (tidak terbatas bulan ini saja)
      // FIX #3: Gunakan fuzzy match agar konsisten dengan tampilan di StudentMaterialsModule
      const _sessionMatchesBadge = (mGroup, sGroup) => {
        if (!mGroup || !sGroup) return false;
        if (mGroup === sGroup) return true;
        const a = mGroup.toLowerCase(), b = sGroup.toLowerCase();
        return a.includes(b) || b.includes(a);
      };
      const allMyMats = (db.materials || []).filter(m => _sessionMatchesBadge(m.sessionGroup, mySession));
      const pendingMats = allMyMats.filter(m => {
         const mySub = (m.submissions || []).find(s => s.studentId === currentUser.studentId);
         return !mySub; // Belum ada pengumpulan tugas
      });
      counts.my_materials = pendingMats.length;

      // 3. Kalender Akademik: Jumlah jadwal kelas yang harus dihadiri HARI INI
      const classesToday = db.calendar.filter(c => c.date === today && (c.sessionGroup || c.name) === mySession);
      counts.calendar = classesToday.length;

      // 4. Announcements: Jumlah pengumuman baru yang dipublish HARI INI
      const annToday = db.announcements.filter(a => a.date === today);
      counts.announcements = annToday.length;

      // 5. Daily Speaking Challenge: Badge merah (1) jika hari ini belum mengerjakan
      const hasDoneChallenge = studentRec?.lastSpeakingChallengeDate === today;
      counts.speaking_challenge = hasDoneChallenge ? 0 : 1;
    }
    
    return counts;
  }, [currentUser, db]);

  if (!currentUser) return <LoginScreen onLogin={handleLogin} isDbLoaded={isDbLoaded} language={language} setLanguage={setLanguage} />;

  // FORCE PASSWORD CHANGE JIKA DIMINTA
  if (currentUser.mustChangePassword) {
    return <ForcePasswordChangeScreen user={currentUser} db={db} setDb={setDb} setCurrentUser={setCurrentUser} showToast={showToast} setSidebarOpen={setSidebarOpen} />;
  }

  // NAVIGATION: Menambahkan akses menu khusus student
  const NAVIGATION = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard, roles: ['admin', 'tutor', 'student'] },
    { id: 'announcements', label: 'Announcements', icon: Bell, roles: ['admin', 'tutor', 'student'] },
    { id: 'students', label: 'Students Directory', icon: Users, roles: ['admin', 'tutor'] },
    { id: 'tutors', label: 'Tutors Directory', icon: Briefcase, roles: ['admin'] },
    { id: 'calendar', label: 'Academic Calendar', icon: CalendarIcon, roles: ['admin', 'tutor', 'student'] },
    { id: 'tutor_attendance', label: 'Tutor Check-In', icon: Activity, roles: ['admin', 'tutor'] },
    { id: 'student_attendance', label: 'Student Attendance', icon: UserCheck, roles: ['admin', 'tutor'] },
    { id: 'my_attendance', label: 'My Attendance', icon: UserCheck, roles: ['student'] }, // Khusus Student
    { id: 'journals', label: 'Learning Journals', icon: BookOpen, roles: ['admin', 'tutor'] },
    { id: 'my_journals', label: 'My Learning Journal', icon: BookOpen, roles: ['student'] }, // Khusus Student
    { id: 'materials', label: 'Materials & Tasks', icon: LinkIcon, roles: ['admin', 'tutor'] },
    { id: 'my_materials', label: 'My Materials', icon: LinkIcon, roles: ['student'] }, // Khusus Student
    { id: 'speaking_challenge', label: 'Daily Speaking', icon: Mic, roles: ['admin', 'tutor', 'student'] }, // Daily Challenge (All Roles)
    { id: 'my_quests', label: 'My Quests & Badges', icon: Trophy, roles: ['student'] }, // Gamifikasi Khusus Student
    { id: 'assessments', label: 'Monthly Assessment', icon: CheckSquare, roles: ['admin', 'tutor'] },
    { id: 'my_assessments', label: 'My Assessment', icon: Award, roles: ['student'] }, // Khusus Student
    { id: 'payments', label: 'Payments', icon: DollarSign, roles: ['admin'] },
    { id: 'my_payments', label: 'My Payment Status', icon: DollarSign, roles: ['student'] }, // Khusus Student
    { id: 'payroll', label: 'Payroll', icon: FileText, roles: ['admin'] },
    { id: 'history', label: 'History & Reports', icon: BarChart3, roles: ['admin', 'tutor'] },
    { id: 'my_report', label: 'My Academic Report', icon: FileText, roles: ['student'] }, // Rapor Khusus Student
    { id: 'settings', label: 'User Management', icon: Settings, roles: ['admin'] },
    { id: 'account_settings', label: 'My Profile', icon: User, roles: ['tutor', 'student'] }, // Tutor & Student ganti pass
    { id: 'system_logs', label: 'System Logs', icon: Terminal, roles: ['admin'] },
    { id: 'recycle_bin', label: 'Recycle Bin', icon: ArchiveRestore, roles: ['admin'] },
    { id: 'export', label: 'Data Export', icon: Database, roles: ['admin'] },
  ];

  const renderContent = () => {
    // SECURITY: Pastikan tab yang diminta sesuai dengan role user.
    // Nav sidebar sudah memfilter tampilan, tapi tanpa guard ini setActiveTab
    // yang dipanggil dari dashboard/widget bisa membuka tab yang tidak diizinkan.
    const allowedTab = NAVIGATION.find(n => n.id === activeTab);
    if (allowedTab && !allowedTab.roles.includes(currentUser.role)) {
      return <div className="p-8 text-center text-gray-400">Access denied.</div>;
    }
    switch (activeTab) {
      case 'dashboard':
        return <Dashboard db={db} setDb={setDb} user={currentUser} setActiveTab={setActiveTab} isCloudConnected={isCloudConnected} language={language} showToast={showToast} />;
      case 'students':
        return <StudentsModule db={db} setDb={setDb} generateId={generateId} showToast={showToast} softDelete={softDelete} user={currentUser} />;
      case 'tutors':
        return <TutorsModule db={db} setDb={setDb} generateId={generateId} showToast={showToast} softDelete={softDelete} />;
      case 'student_attendance':
        return <StudentAttendanceModule db={db} setDb={setDb} showToast={showToast} softDelete={softDelete} user={currentUser} generateId={generateId} requestConfirm={requestConfirm} setModuleDirty={setModuleDirty} />;
      case 'tutor_attendance':
        return <TutorAttendanceModule db={db} setDb={setDb} user={currentUser} showToast={showToast} softDelete={softDelete} generateId={generateId} />;
      case 'my_attendance': // STUDENT: Read Only Attendance
        return <StudentReadOnlyAttendanceModule db={db} user={currentUser} language={language} />;
      case 'journals':
        return <JournalsModule db={db} setDb={setDb} user={currentUser} showToast={showToast} generateId={generateId} softDelete={softDelete} />;
      case 'my_journals': // STUDENT: Read Only Journal
        return <StudentReadOnlyJournalsModule db={db} user={currentUser} language={language} />;
      case 'assessments':
        return <AssessmentsModule db={db} setDb={setDb} generateId={generateId} showToast={showToast} user={currentUser} />;
      case 'my_assessments': // STUDENT: Read Only Assessment
        return <StudentReadOnlyAssessmentModule db={db} user={currentUser} language={language} />;
      case 'materials': // TUTOR & ADMIN: Manage Materials & Tasks
        return <MaterialsModule db={db} setDb={setDb} generateId={generateId} showToast={showToast} softDelete={softDelete} user={currentUser} />;
      case 'my_materials': // STUDENT: View Materials & Submit Comment
        return <StudentMaterialsModule db={db} setDb={setDb} user={currentUser} showToast={showToast} language={language} />;
      case 'speaking_challenge': // STUDENT: Daily Speaking Challenge Menu
        return <StudentSpeakingChallengeModule db={db} setDb={setDb} user={currentUser} showToast={showToast} language={language} />;
      case 'my_quests': // STUDENT: Gamification Quests
        return <StudentQuestsModule db={db} user={currentUser} language={language} setActiveTab={setActiveTab} downloadPNG={downloadPNG} handleShareImage={handleShareImage} />;
      case 'my_report': // STUDENT: Academic Report
        return <StudentReadOnlyReportModule db={db} user={currentUser} downloadPNG={downloadPNG} handleShareImage={handleShareImage} language={language} />;
      case 'payments':
        return (
          <PaymentsModule
            db={db}
            setDb={setDb}
            generateId={generateId}
            showToast={showToast}
            handlePrint={() => window.print()}
            handleShareImage={handleShareImage}
            downloadPNG={downloadPNG}
            softDelete={softDelete}
            language={language}
            isDbDirty={isDbDirty}
        />
      );
      case 'my_payments': // STUDENT: Read Only Payments
        return <StudentReadOnlyPaymentModule db={db} user={currentUser} downloadPNG={downloadPNG} handleShareImage={handleShareImage} language={language} showToast={showToast} />;
      case 'payroll':
        return (
          <PayrollModule
            db={db}
            setDb={setDb}
            generateId={generateId}
            showToast={showToast}
            handlePrint={() => window.print()}
            handleShareImage={handleShareImage}
            downloadPNG={downloadPNG}
            softDelete={softDelete}
            requestConfirm={requestConfirm}
          />
        );
      case 'calendar':
        return <CalendarModule db={db} setDb={setDb} generateId={generateId} user={currentUser} showToast={showToast} softDelete={softDelete} />;
      case 'announcements':
        return <AnnouncementsModule db={db} setDb={setDb} generateId={generateId} user={currentUser} showToast={showToast} softDelete={softDelete} setActiveTab={setActiveTab} />;
      case 'history':
        return <HistoryReportsModule db={db} setDb={setDb} showToast={showToast} handlePrint={() => window.print()} user={currentUser} handleShareImage={handleShareImage} downloadPNG={downloadPNG} />;
      case 'settings':
        return <SettingsModule db={db} setDb={setDb} generateId={generateId} user={currentUser} showToast={showToast} requestConfirm={requestConfirm} getAuthToken={getAuthToken} dbVersion={dbVersion} currentUser={currentUser} setSyncStatus={setSyncStatus} />;
      case 'account_settings':
        return <AccountSettingsModule db={db} setDb={setDb} user={currentUser} setCurrentUser={setCurrentUser} showToast={showToast} language={language} />;
      case 'system_logs':
        return <SystemLogsModule logs={logs} setLogs={setLogs} showToast={showToast} requestConfirm={requestConfirm} />;
      case 'recycle_bin':
        return <RecycleBinModule db={db} setDb={setDb} showToast={showToast} requestConfirm={requestConfirm} />;
      case 'export':
        return <DataExportModule db={db} />;
      default:
        return <div className="p-8 text-center text-gray-400">Module under construction</div>;
    }
  };

  return (
    <div className="flex h-screen bg-[#0B0F19] text-[#F3F4F6] font-['Poppins',sans-serif] overflow-hidden selection:bg-[#00D4FF] selection:text-[#0B0F19]">
      <Toaster
        position="top-right"
        toastOptions={{
          style: {
            background: '#151B26',
            color: '#F3F4F6',
            border: '1px solid #374151',
          },
        }}
      />

      {toast && (
        <div key={toast.id} className={`fixed top-4 right-4 z-50 px-4 py-3 rounded-lg shadow-2xl flex items-center gap-2 animate-bounce print:hidden ${
            toast.type === 'error' ? 'bg-red-500' : toast.type === 'warning' ? 'bg-yellow-500 text-[#0B0F19]' : 'bg-[#00D4FF] text-[#0B0F19] font-semibold'
        }`}>
          {toast.type === 'error' ? <XCircle size={20} /> : toast.type === 'warning' ? <AlertCircle size={20} /> : <CheckCircle2 size={20} />}
          {toast.msg}
        </div>
      )}

      <CustomModal isOpen={!!logoutConfirm} onClose={() => setLogoutConfirm(false)} title={language === 'id' ? 'Keluar' : 'Log Out'}>
        <div className="flex items-center gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-red-500/10 flex items-center justify-center border border-red-500/20 flex-shrink-0">
            <LogOut size={20} className="text-red-400" />
          </div>
          <p className="text-gray-400">{language === 'id' ? 'Apakah Anda yakin ingin keluar?' : 'Are you sure you want to log out?'}</p>
        </div>
        <div className="flex justify-end gap-3">
          <Button variant="ghost" onClick={() => setLogoutConfirm(false)}>{language === 'id' ? 'Cancel' : 'Cancel'}</Button>
          <Button className="bg-red-500 hover:bg-red-600 text-white border-none shadow-none" onClick={confirmLogout}>{language === 'id' ? 'Konfirmasi' : 'Confirm'}</Button>
        </div>
      </CustomModal>

      <CustomModal isOpen={!!confirmDialog} onClose={() => setConfirmDialog(null)} title={confirmDialog?.title || ''} zIndexClass="z-[60]">
        <p className="text-gray-400 mb-6">{confirmDialog?.message}</p>
        <div className="flex justify-end gap-3">
          <Button variant="ghost" onClick={() => setConfirmDialog(null)}>Cancel</Button>
          <Button className="bg-red-500 hover:bg-red-600 text-white border-none shadow-none" onClick={confirmDialog?.onConfirm}>Confirm</Button>
        </div>
      </CustomModal>

      <div className={`fixed inset-y-0 left-0 z-40 w-64 bg-[#0A0E17] border-r border-gray-800 transform transition-transform duration-300 ease-in-out ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full'} md:relative md:translate-x-0 print:hidden flex flex-col`}>
        <div className="p-6 flex items-center justify-between">
          <div>
            <img src={LOGO_URL} alt="Logo" className="h-10 w-auto mb-2 opacity-90 drop-shadow-[0_0_8px_rgba(0,212,255,0.3)]" />
            <h1 className="text-xl font-bold text-[#00D4FF] leading-tight">ECG Academic<br />Suite</h1>
          </div>
          <button onClick={() => setSidebarOpen(false)} className="md:hidden text-gray-400 hover:text-white"><X size={24} /></button>
        </div>

        <div className="flex-1 overflow-y-auto py-4 px-3 space-y-1 custom-scrollbar">
          {/* BUG FIX #10: getNavLabel moved outside map() — was being re-created on every render */}
          {(() => {
            const NAV_LABEL_DICT_ID = {
              'Dashboard': 'Beranda',
              'Announcements': 'Pengumuman',
              'Students Directory': 'Direktori Siswa',
              'Tutors Directory': 'Direktori Tutor',
              'Academic Calendar': 'Kalender Akademik',
              'Tutor Check-In': 'Absensi Tutor',
              'Student Attendance': 'Kehadiran Siswa',
              'My Attendance': 'Kehadiran Saya',
              'Learning Journals': 'Jurnal Belajar',
              'My Learning Journal': 'Jurnal Belajar',
              'Monthly Assessment': 'Penilaian Bulanan',
              'My Assessment': 'Penilaian Saya',
              'Materials & Tasks': 'Materi & Tugas',
              'My Materials': 'Materi Saya',
              'Daily Speaking': 'Tantangan Harian',
              'My Quests & Badges': 'Misi & Lencana',
              'My Academic Report': 'Rapor Akademik',
              'Payments': 'Pembayaran (SPP)',
              'My Payment Status': 'Status Pembayaran',
              'Payroll': 'Penggajian (Gaji)',
              'History & Reports': 'Riwayat & Laporan',
              'User Management': 'Manajemen Pengguna',
              'My Profile': 'Profil Saya',
              'System Logs': 'Log Sistem',
              'Recycle Bin': 'Tempat Sampah'
            };
            const getNavLabel = (label) => language === 'en' ? label : (NAV_LABEL_DICT_ID[label] || label);
            return NAVIGATION.filter((nav) => nav.roles.includes(currentUser.role)).map((nav) => {
            const badge = badgeCounts[nav.id] || 0;

            return (
              <button
                key={nav.id}
                onClick={() => {
                  if (moduleDirtyRef.current && nav.id !== activeTab) {
                    if (!window.confirm('You have unsaved attendance data. Leave this page?\n\nUnsaved changes will be lost.')) return;
                    moduleDirtyRef.current = false;
                  }
                  setActiveTab(nav.id);
                  if (window.innerWidth < 768) setSidebarOpen(false);
                }}
                className={`w-full flex items-center justify-between px-4 py-3 rounded-lg transition-colors ${
                  activeTab === nav.id ? 'bg-[#151B26] text-[#00D4FF] border-l-2 border-[#00D4FF] shadow-[inset_0_0_15px_rgba(0,212,255,0.05)]' : 'text-gray-400 hover:bg-[#151B26] hover:text-white'
                }`}
              >
                <div className="flex items-center gap-3">
                  <nav.icon size={20} className={activeTab === nav.id ? 'text-[#00D4FF]' : ''} />
                  <span className="font-medium text-sm">{getNavLabel(nav.label)}</span>
                </div>
                {badge > 0 && (
                  <span className="bg-red-500 text-white text-[11px] font-bold px-2 py-0.5 rounded-full shadow-[0_0_10px_rgba(239,68,68,0.4)] min-w-[20px] text-center">
                    {badge}
                  </span>
                )}
              </button>
            );
          });
          })()}
        </div>

        <div className="p-4 border-t border-gray-800 bg-[#0B0F19]">
             <div className="flex items-center justify-between mb-4 bg-[#151B26] p-2 rounded-lg border border-gray-800">
                <span className="text-xs text-gray-400 font-semibold px-2">{language === 'id' ? 'Bahasa Pilihan' : 'Preferred Language'}</span>
                <div className="flex bg-[#0B0F19] rounded-md border border-gray-700 overflow-hidden">
                   <button onClick={() => setLanguage('en')} className={`px-2.5 py-1 text-[11px] font-bold transition-colors ${language === 'en' ? 'bg-[#00D4FF] text-[#0B0F19]' : 'text-gray-500 hover:text-white'}`}>EN</button>
                   <button onClick={() => setLanguage('id')} className={`px-2.5 py-1 text-[11px] font-bold transition-colors ${language === 'id' ? 'bg-[#00D4FF] text-[#0B0F19]' : 'text-gray-500 hover:text-white'}`}>ID</button>
                </div>
             </div>
          <Button variant="ghost" className="w-full justify-start text-blue-400 hover:text-blue-300 hover:bg-blue-500/10 transition-colors mb-2" icon={RefreshCw} onClick={refreshBeforeEdit}>
            {language === 'id' ? 'Sinkronkan Data' : 'Sync Data'}
          </Button>
          
          <div className="relative mb-4">
             <button 
                onClick={() => setShowAccountMenu(!showAccountMenu)} 
                className="flex items-center gap-3 w-full p-2.5 hover:bg-[#151B26] rounded-xl transition-all border border-transparent hover:border-gray-800"
             >
                <div className="w-9 h-9 rounded-full bg-gradient-to-br from-[#00D4FF] to-[#2563EB] flex items-center justify-center text-white font-bold shadow-lg shadow-[#00D4FF]/20 shrink-0">
                  {currentUser.name.charAt(0)}
                </div>
                <div className="overflow-hidden text-left flex-1">
                  <p className="text-sm font-bold text-white truncate">{currentUser.name}</p>
                  <p className="text-[11px] text-[#00D4FF] uppercase tracking-wider font-semibold truncate">{currentUser.role}</p>
                </div>
                <ChevronDown size={16} className={`text-gray-400 shrink-0 transition-transform ${showAccountMenu ? 'rotate-180' : ''}`} />
             </button>

             {showAccountMenu && (
                <div className="absolute bottom-full left-0 w-full bg-[#151B26] border border-gray-700 rounded-xl shadow-2xl mb-2 overflow-hidden z-50 animation-fade-in">
                   <div className="p-3 border-b border-gray-800 flex justify-between items-center">
                      <span className="text-[11px] font-semibold text-gray-400 uppercase tracking-widest">{language === 'id' ? 'Ganti Akun' : 'Switch Account'}</span>
                   </div>
                   <div className="max-h-48 overflow-y-auto custom-scrollbar p-1">
                      {savedAccounts.map((acc, i) => (
                         <div key={i} className={`flex items-center w-full rounded-lg mb-1 relative group ${currentUser.username === acc.username ? 'bg-[#00D4FF]/10' : 'hover:bg-[#0B0F19]'}`}>
                            <button onClick={() => switchAccount(acc)} className="flex-1 flex items-center gap-3 p-2.5 text-left truncate">
                               <div className="w-7 h-7 rounded-full bg-gradient-to-br from-[#00D4FF] to-[#2563EB] flex items-center justify-center text-white font-bold text-xs shrink-0">
                                  {acc.name.charAt(0)}
                               </div>
                               <div className="overflow-hidden flex-1">
                                  <p className="text-xs font-bold text-white truncate">{acc.name}</p>
                                  <p className="text-[11px] text-gray-400 uppercase tracking-wider truncate">{acc.role}</p>
                               </div>
                            </button>
                            {currentUser.username === acc.username ? (
                               <CheckCircle2 size={16} className="text-[#00D4FF] shrink-0 mr-3" />
                            ) : (
                               <button onClick={(e) => removeSavedAccount(e, acc.username)} className="p-2 text-gray-500 hover:text-red-400 opacity-100 md:opacity-0 md:group-hover:opacity-100 transition-opacity absolute right-1">
                                  <X size={14} />
                               </button>
                            )}
                         </div>
                      ))}
                   </div>
                   <div className="p-2 border-t border-gray-800 bg-[#0A0E17]">
                      <button onClick={() => { setCurrentUser(null); sessionStorage.removeItem('ecg_active_session'); }} className="flex items-center justify-center gap-2 w-full p-2 text-xs text-blue-400 hover:bg-blue-500/10 rounded-lg transition-colors font-bold">
                         <Plus size={14} /> {language === 'id' ? 'Tambah Akun Lain' : 'Add Another Account'}
                      </button>
                   </div>
                </div>
             )}
          </div>

          <Button variant="ghost" className="w-full justify-start text-red-400 hover:text-red-300 hover:bg-red-500/10 transition-colors" icon={LogOut} onClick={handleLogout}>
            {language === 'id' ? 'Keluar' : 'Sign Out'}
          </Button>
        </div>
      </div>

      <main className="flex-1 overflow-y-auto p-4 md:p-8 pt-20 md:pt-12 relative custom-scrollbar">
          
          {/* DESKTOP CLOUD AUTO-SAVE INDICATOR */}
          <div className="hidden md:flex absolute top-6 right-8 z-50 print:hidden pointer-events-none">
             <CloudAutoSaveIndicator status={syncStatus} language={language} />
          </div>

          {/* MOBILE STICKY HEADER (Hamburger Menu & Back Button) */}
          <div className="fixed top-0 left-0 right-0 z-30 bg-[#0B0F19]/95 backdrop-blur-xl border-b border-gray-800/80 px-4 h-16 flex items-center justify-between md:hidden print:hidden shadow-md">
             {activeTab === 'dashboard' ? (
                <button onClick={() => setSidebarOpen(true)} className="p-2 -ml-2 text-gray-300 hover:text-[#00D4FF] transition-colors flex items-center gap-3">
                   <Menu size={24} />
                   <span className="font-bold text-sm text-[#00D4FF] tracking-wide uppercase">{language === 'id' ? 'Menu Utama' : 'Main Menu'}</span>
                </button>
             ) : (
                <button onClick={() => setActiveTab('dashboard')} className="p-2 -ml-2 text-gray-400 hover:text-white transition-colors flex items-center gap-2">
                   <ArrowLeft size={22} />
                   <span className="text-sm font-bold tracking-wide">{language === 'id' ? 'Kembali' : 'Back'}</span>
                </button>
             )}
             
             {/* KANAN: Sync Indicator & Refresh Button */}
             <div className="flex items-center gap-2">
                <CloudAutoSaveIndicator status={syncStatus} language={language} />
                <button onClick={refreshBeforeEdit} className="p-2 text-blue-400 hover:text-blue-300 transition-colors" title="Sync Data">
                   <RefreshCw size={20} />
                </button>
             </div>
          </div>

          {/* DESKTOP BACK BUTTON (Only when not on dashboard) */}
          {activeTab !== 'dashboard' && (
            <div className="hidden md:flex sticky top-[-32px] z-40 bg-[#0B0F19]/95 backdrop-blur-md pt-8 pb-4 -mt-8 mb-6 border-b border-gray-800/80 print:hidden w-full max-w-7xl mx-auto justify-between items-center">
              <Button
                variant="secondary"
                onClick={() => setActiveTab('dashboard')}
                icon={ArrowLeft}
                className="w-auto min-h-[44px] flex items-center justify-center bg-[#151B26] hover:bg-[#1E293B] border border-gray-700 hover:border-[#00D4FF]/60 text-gray-200 hover:text-white transition-all shadow-md rounded-xl px-6 font-semibold tracking-wide"
              >
                {language === 'id' ? 'Kembali ke Beranda' : 'Back to Dashboard'}
              </Button>
              <Button onClick={refreshBeforeEdit} variant="secondary" icon={RefreshCw} className="border-blue-500/30 hover:border-blue-500 hover:bg-blue-500/10 text-blue-400 shadow-md">
                {language === 'id' ? 'Sinkronkan Data' : 'Sync Latest Data'}
              </Button>
            </div>
          )}

          <div className="max-w-7xl mx-auto space-y-6 print:max-w-full print:space-y-0">
            {renderContent()}
          </div>
        </main>
      </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <MainApp />
    </ErrorBoundary>
  );
}

function StudentsModule({ db, setDb, generateId, showToast, softDelete, user }) {
  const validLevelsForTutor = user?.role === 'tutor' 
    ? LEVELS.filter(lvl => (CLASS_MAPPING[lvl] || []).some(cls => parseSessions(user.teachingSession).includes(getSessionGroup(cls))))
    : LEVELS;
  const defaultLevel = validLevelsForTutor.length > 0 ? validLevelsForTutor[0] : LEVELS[0];
  const defaultClass = CLASS_MAPPING[defaultLevel].find(cls => user?.role === 'tutor' ? parseSessions(user.teachingSession).includes(getSessionGroup(cls)) : true) || CLASS_MAPPING[defaultLevel][0];
  
  const [formData, setFormData] = useState({ id: '', name: '', gender: 'Male', level: defaultLevel, class: defaultClass, paymentPlan: 'Monthly', status: 'Active', teacherComment: '', sessionOverride: 'Default', enrollmentStatus: 'Returning', whatsapp: '', billingStartMonth: '' });
  const [isAdding, setIsAdding] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearchTerm = useDebounce(searchTerm, 300);
  const [filterLevel, setFilterLevel] = useState('');
  const [filterSession, setFilterSession] = useState('');
  const [filterClass, setFilterClass] = useState('');
  const [filterPaymentPlan, setFilterPaymentPlan] = useState('');

  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  // Hitung kehadiran bulan berjalan per siswa (Present saja)
  // Fix 4: Only count valid-session attendance for overridden students
  const attendanceThisMonth = useMemo(() => {
    const now = new Date();
    const monthPrefix = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const map: Record<string, number> = {};
    (db.studentAttendance || []).forEach(a => {
      if (String(a.date || '').startsWith(monthPrefix) && a.status === 'Present') {
        const studentObj = db.students.find(s => s.id === a.studentId);
        if (isAttendanceValidForStudent(a, studentObj)) {
          map[a.studentId] = (map[a.studentId] || 0) + 1;
        }
      }
    });
    return map;
  }, [db.studentAttendance, db.students]);

  useEffect(() => {
    if (!formData.id && formData.level) {
      const availableClasses = CLASS_MAPPING[formData.level] || [];
      const validClasses = user?.role === 'tutor' 
         ? availableClasses.filter(cls => parseSessions(user.teachingSession).includes(getSessionGroup(cls)))
         : availableClasses;
      setFormData(prev => ({ ...prev, class: validClasses[0] || availableClasses[0] }));
    }
  }, [formData.level, user]);

  useEffect(() => {
    setCurrentPage(1);
  }, [debouncedSearchTerm, filterLevel, filterSession, filterClass, filterPaymentPlan, rowsPerPage]);

  const handleSave = (e) => {
    e.preventDefault();
    // Simpan langsung dalam format internasional (628...) agar siap dipakai untuk link wa.me tanpa konversi lagi.
    // Format ini tidak diawali angka 0, jadi otomatis aman dari masalah auto-konversi angka di Google Sheets.
    const finalWa = normalizeWhatsapp(formData.whatsapp);
    // FIX BUG #2: tambah updatedAt agar LWW di mergeByIds tahu versi lokal lebih baru dari cloud
    const rec = { ...formData, whatsapp: finalWa, id: formData.id || generateId('STU', 'students'), updatedAt: getSyncTimestamp() };
    setDb((prev) => ({ 
      ...prev, 
      students: formData.id 
        ? prev.students.map((s) => {
            if (s.id !== formData.id) return s;
            // BUGFIX SESSION OVERRIDE: Merge existing student record dengan form data agar field
            // yang tidak ada di form (bonusExp, speakingChallengeCompletedCount, dll) tidak hilang.
            // Ini juga mencegah sessionOverride hilang jika versi lama student dari cloud tidak punya field ini.
            return { ...s, ...rec };
          })
        : [...prev.students, rec]
    }));
    showToast('Student saved');
    setIsAdding(false);
  };

  const filtered = sortStudentsLogically(db.students.filter((s) => {
    const matchSearch = (s.name || '').toLowerCase().includes((debouncedSearchTerm || '').toLowerCase());
    const matchLevel = filterLevel ? s.level === filterLevel : true;
    const matchClass = filterClass ? s.class === filterClass : true;
    // FIX: Untuk tutor, filterSession hanya boleh memilih dari sesi miliknya sendiri
    // Jika filterSession kosong, tampilkan semua siswa dari semua sesi tutor (matchTutor sudah menjaga scope)
    const matchSessionFilter = filterSession ? getStudentSession(s) === filterSession : true;
    const matchTutor = user?.role === 'tutor' ? parseSessions(user.teachingSession).includes(getStudentSession(s)) : true;
    const matchPlan = filterPaymentPlan ? s.paymentPlan === filterPaymentPlan : true;
    return matchSearch && matchLevel && matchClass && matchSessionFilter && matchTutor && matchPlan;
  }));

  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filtered.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filtered.length / (rowsNum || 1));
  const paginatedData = isAll ? filtered : filtered.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white mb-1">Students Directory</h2><p className="text-gray-400 text-sm">Manage student records, enrollment, and class placement.</p></div>
        {(!user || user.role === 'admin' || user.role === 'tutor') && (
          <Button onClick={() => { setFormData({ id: '', name: '', gender: 'Male', level: defaultLevel, class: defaultClass, paymentPlan: 'Monthly', status: 'Active', teacherComment: '', sessionOverride: 'Default', enrollmentStatus: 'Returning', whatsapp: '', billingStartMonth: '' }); setIsAdding(!isAdding); }} icon={Plus}>Add Student</Button>
        )}
      </div>
      {isAdding && (
        <Card className="border border-emerald-500/20 shadow-[0_0_24px_rgba(16,185,129,0.06)]">
          <div className="flex items-center gap-3 mb-6 pb-4 border-b border-gray-800">
            <div className="p-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/20">
              <Users size={20} className="text-emerald-400" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white">{formData.id ? 'Edit Student' : 'Add New Student'}</h3>
              <p className="text-xs text-gray-500 mt-0.5">Fields marked <span className="text-red-400">*</span> are required</p>
            </div>
          </div>
          <form onSubmit={handleSave}>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-1">
              <Input label="Full Name" value={formData.name} onChange={(v) => setFormData({ ...formData, name: v })} required />
              <div>
                <Input label="No. WhatsApp (Optional)" value={formData.whatsapp} onChange={(v) => setFormData({ ...formData, whatsapp: v })} placeholder="08xxx or 628xxx" />
                <p className="text-[11px] text-gray-500 mt-1 mb-3 px-1 leading-tight">Format 08... or 628... — auto-saved as international (628...)</p>
              </div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <Input label="Gender" type="select" options={['Male', 'Female']} value={formData.gender} onChange={(v) => setFormData({ ...formData, gender: v })} required />
              <Input label="Enrollment Status" type="select" options={['New', 'Returning']} value={formData.enrollmentStatus} onChange={(v) => setFormData({ ...formData, enrollmentStatus: v })} required />
            </div>
            {/* BILLING START MONTH — hanya tampil jika siswa baru atau admin mengatur */}
            <div className="mb-4 p-3 rounded-lg border border-amber-500/20 bg-amber-500/5">
              <label className="block text-sm text-amber-400 font-semibold mb-1 flex items-center gap-2">
                <CalendarIcon size={14} /> Billing Start Month <span className="text-gray-500 font-normal">(Opsional)</span>
              </label>
              <input
                type="month"
                className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-amber-400 transition-all"
                value={formData.billingStartMonth || ''}
                onChange={(e) => setFormData({ ...formData, billingStartMonth: e.target.value })}
              />
              <p className="text-[11px] text-gray-500 mt-1.5 leading-tight">
                Kosongkan jika tagihan mulai bulan ini. Isi jika siswa baru akan mulai membayar bulan depan atau seterusnya — bulan sebelum tanggal ini <strong className="text-amber-400">tidak akan dihitung</strong> dalam expected collected.
              </p>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <Input label="Level" type="select" options={validLevelsForTutor} value={formData.level} onChange={(v) => setFormData({ ...formData, level: v })} required />
              <Input label="Class" type="select" options={(CLASS_MAPPING[formData.level] || []).filter(cls => user?.role === 'tutor' ? parseSessions(user.teachingSession).includes(getSessionGroup(cls)) : true)} value={formData.class} onChange={(v) => setFormData({ ...formData, class: v })} required />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <Input label="Payment Plan" type="select" options={['Monthly', 'Per Visit', 'Free']} value={formData.paymentPlan} onChange={(v) => setFormData({ ...formData, paymentPlan: v })} required />
              <Input label="Status" type="select" options={['Active', 'Inactive']} value={formData.status} onChange={(v) => setFormData({ ...formData, status: v })} />
            </div>
            <div className="mb-6">
              <Input label="Session Override (Optional)" type="select" options={['Default', ...SESSIONS]} value={formData.sessionOverride || 'Default'} onChange={(v) => setFormData({ ...formData, sessionOverride: v })} />
              <p className="text-[11px] text-gray-500 mt-1 mb-2 px-1 leading-tight">Leave "Default" to follow the class session, or pick another if this student joins a different time slot.</p>
            </div>
            <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
              <Button variant="secondary" onClick={() => setIsAdding(false)}>Cancel</Button>
              <Button type="submit" icon={formData.id ? Edit2 : Plus}>{formData.id ? 'Update Student' : 'Save Student'}</Button>
            </div>
          </form>
        </Card>
      )}
      <Card className="p-0 overflow-hidden flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 space-y-3">
          {/* Row 1: Search bar full width */}
          <div className="relative">
            <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none" />
            <input
              type="text"
              placeholder="Search students by name or ID..."
              className="w-full bg-[#151B26] border border-gray-700 rounded-lg pl-10 pr-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all placeholder-gray-600"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          {/* Row 2: 4 filter dropdowns evenly distributed */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all text-sm" value={filterLevel} onChange={(e) => { setFilterLevel(e.target.value); setFilterClass(''); }}>
              <option value="">All Levels</option>
              {validLevelsForTutor.map((l) => <option key={l} value={l}>{l}</option>)}
            </select>
            <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all text-sm disabled:opacity-40 disabled:cursor-not-allowed" value={filterClass} onChange={(e) => setFilterClass(e.target.value)} disabled={!filterLevel}>
              <option value="">All Classes</option>
              {filterLevel && (CLASS_MAPPING[filterLevel] || []).filter(cls => user?.role === 'tutor' ? parseSessions(user.teachingSession).includes(getSessionGroup(cls)) : true).map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <select
              className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all text-sm"
              value={filterSession}
              onChange={(e) => setFilterSession(e.target.value)}
            >
              {user?.role === 'tutor' ? (
                <>
                  <option value="">All My Sessions</option>
                  {parseSessions(user.teachingSession).map((s) => <option key={s} value={s}>{s}</option>)}
                </>
              ) : (
                <>
                  <option value="">All Sessions</option>
                  {SESSIONS.map((l) => <option key={l} value={l}>{l}</option>)}
                </>
              )}
            </select>
            <select
              className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all text-sm"
              value={filterPaymentPlan}
              onChange={(e) => setFilterPaymentPlan(e.target.value)}
            >
              <option value="">All Plans</option>
              <option value="Monthly">Monthly</option>
              <option value="Per Visit">Per Visit</option>
              <option value="Free">Free</option>
            </select>
          </div>
        </div>
        {/* ── MOBILE CARD LIST (< sm) ── */}
        <div className="sm:hidden divide-y divide-gray-800">
          {paginatedData.length === 0 ? (
            <EmptyState icon={Users} title="No students found" description="Try adjusting your search or filter criteria." />
          ) : paginatedData.map((s, index) => {
            const count = attendanceThisMonth[s.id] || 0;
            const isPerVisit = s.paymentPlan === 'Per Visit';
            return (
              <div key={s.id} className="p-4 space-y-3 hover:bg-[#0B0F19]/60 transition-colors">
                {/* Row 1: number + name + status */}
                <div className="flex items-start justify-between gap-2">
                  <div className="flex items-center gap-2 min-w-0">
                    <span className="shrink-0 inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF]">{String(startIndex + index + 1).padStart(2,'0')}</span>
                    <div className="min-w-0">
                      <div className="flex flex-wrap items-center gap-1">
                        <span className="font-bold text-white text-sm truncate">{s.name}</span>
                        <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} />
                      </div>
                      <p className="font-mono text-[11px] text-gray-500">{s.id}</p>
                    </div>
                  </div>
                  <Badge status={s.status} />
                </div>
                {/* Row 2: level / class / plan / attendance */}
                <div className="flex flex-wrap gap-2 text-xs">
                  <span className="text-[#00D4FF] font-semibold">{s.level}</span>
                  <span className="text-gray-400">·</span>
                  <span className="font-bold text-gray-300">{s.class}</span>
                  <span className={`px-2 py-0.5 rounded font-semibold uppercase tracking-wide border ${s.paymentPlan === 'Per Visit' ? 'bg-purple-500/20 text-purple-300 border-purple-500/50' : s.paymentPlan === 'Free' ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/50' : 'bg-blue-500/20 text-blue-300 border-blue-500/50'}`}>{s.paymentPlan}</span>
                  <span className={`inline-flex items-center justify-center px-2 py-0.5 rounded-full font-bold ${count === 0 ? 'bg-gray-700/50 text-gray-500' : isPerVisit ? 'bg-purple-500/20 text-purple-300 border border-purple-500/40' : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'}`}>{count}x hadir</span>
                </div>
                {/* Row 3: WA + actions */}
                {(!user || user.role === 'admin' || user.role === 'tutor') && (
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-gray-500 text-xs truncate">{s.whatsapp ? String(s.whatsapp).replace(/^'/,'') : <span className="italic text-gray-600">No WA</span>}</span>
                    <div className="flex gap-1 shrink-0">
                      {s.whatsapp ? (
                        <a href={`https://wa.me/${normalizeWhatsapp(s.whatsapp)}`} target="_blank" rel="noopener noreferrer" className="text-green-400 p-2 hover:bg-green-500/10 rounded-lg transition-colors" title="Chat WhatsApp"><MessageCircle size={16}/></a>
                      ) : (
                        <span className="text-gray-700 p-2 cursor-not-allowed"><MessageCircle size={16}/></span>
                      )}
                      <button onClick={() => { setFormData({...s, whatsapp: String(s.whatsapp || '').replace(/^'/,'')}); setIsAdding(true); const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50); }} className="text-blue-400 p-2 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Student"><Edit2 size={16}/></button>
                      {(!user || user.role === 'admin') && (
                        <button onClick={() => softDelete('students', s.id, s.name)} className="text-red-400 p-2 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Student"><Trash2 size={16}/></button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* ── DESKTOP TABLE (≥ sm) ── */}
        <div className="hidden sm:block overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4">ID</th><th className="p-4">Name</th><th className="p-4">Level</th><th className="p-4">Class & Plan</th><th className="p-4 text-center">Attended<br/><span className="text-[10px] font-normal text-gray-500 normal-case tracking-normal">This Month</span></th><th className="p-4">WhatsApp</th><th className="p-4 text-center">Status</th>{(!user || user.role === 'admin' || user.role === 'tutor') && <th className="p-4 text-center">Actions</th>}</tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {paginatedData.map((s, index) => (
              <tr key={s.id} className="hover:bg-[#0B0F19]">
                <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                <td className="p-4 font-mono text-gray-400">{s.id}</td>
                <td className="p-4 text-white font-medium">
                  <div className="flex items-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /></div>
                </td>
                <td className="p-4 text-[#00D4FF]">{s.level}</td>
                <td className="p-4 leading-tight">
                  <span className="font-bold text-gray-300">{s.class}</span><br/>
                  <span className={`text-[11px] px-2 py-0.5 rounded font-semibold uppercase tracking-wide mt-1 inline-block border ${s.paymentPlan === 'Per Visit' ? 'bg-purple-500/20 text-purple-300 border-purple-500/50' : s.paymentPlan === 'Free' ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/50' : 'bg-blue-500/20 text-blue-300 border-blue-500/50'}`}>{s.paymentPlan}</span>
                </td>
                <td className="p-4 text-center">
                  {(() => {
                    const count = attendanceThisMonth[s.id] || 0;
                    const isPerVisit = s.paymentPlan === 'Per Visit';
                    return (
                      <span className={`inline-flex items-center justify-center min-w-[2rem] px-2 py-1 rounded-full text-sm font-bold ${
                        count === 0 ? 'bg-gray-700/50 text-gray-500' :
                        isPerVisit ? 'bg-purple-500/20 text-purple-300 border border-purple-500/40' :
                        'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      }`}>
                        {count}x
                      </span>
                    );
                  })()}
                </td>
                <td className="p-4 text-center"><Badge status={s.status} /></td>
                <td className="p-4 text-gray-400 text-sm">{s.whatsapp ? String(s.whatsapp).replace(/^'/, '') : <span className="text-gray-600 italic text-xs">No WA</span>}</td>
                {(!user || user.role === 'admin' || user.role === 'tutor') && (
                  <td className="p-4 text-center flex justify-center gap-2">
                    {s.whatsapp ? (
                      <a href={`https://wa.me/${normalizeWhatsapp(s.whatsapp)}`} target="_blank" rel="noopener noreferrer" className="text-green-400 p-2.5 hover:bg-green-500/10 rounded-lg transition-colors" title="Chat WhatsApp"><MessageCircle size={18}/></a>
                    ) : (
                      <span className="text-gray-700 p-2.5 cursor-not-allowed" title="No WhatsApp number"><MessageCircle size={18}/></span>
                    )}
                    <button onClick={() => { setFormData({...s, whatsapp: String(s.whatsapp || '').replace(/^'/, '')}); setIsAdding(true); const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Student"><Edit2 size={18} /></button>
                    {(!user || user.role === 'admin') && (
                      <button onClick={() => softDelete('students', s.id, s.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Student"><Trash2 size={18} /></button>
                    )}
                  </td>
                )}
              </tr>
            ))}
            {paginatedData.length === 0 && (
               <tr><td colSpan={9}><EmptyState icon={Users} title="No students found" description="Try adjusting your search or filter criteria." /></td></tr>
            )}
          </tbody>
        </table>
        </div>
        
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
           <div className="flex items-center gap-2">
             <span>Show</span>
             <select 
               value={rowsPerPage} 
               onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} 
               className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer"
             >
               <option value={10}>10</option>
               <option value={20}>20</option>
               <option value={30}>30</option>
               <option value="All">All</option>
             </select>
             <span>entries {filtered.length > 0 && `(Total: ${filtered.length})`}</span>
           </div>
           
           {!isAll && totalPages > 1 && (
             <div className="flex items-center gap-2">
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
               <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
             </div>
           )}
        </div>
      </Card>
    </div>
  );
}

function StudentAttendanceModule({ db, setDb, showToast, softDelete, user, generateId, requestConfirm, setModuleDirty }) {
  const [selectedScheduleId, setSelectedScheduleId] = useState('');
  const [attendanceData, setAttendanceData] = useState({});
  const [editingId, setEditingId] = useState(null);
  const [editStatus, setEditStatus] = useState('');
  const [viewDate, setViewDate] = useState(getTodayDateLocal());
  const [editScheduleIdFilter, setEditScheduleIdFilter] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  // Fix 10: Search by name in Edit Records
  const [editSearchName, setEditSearchName] = useState('');
  // Month/year pre-filter for Edit Records dropdown (keeps list short)
  const [editFilterMonth, setEditFilterMonth] = useState<number>(() => new Date().getMonth() + 1);
  const [editFilterYear, setEditFilterYear] = useState<number>(() => new Date().getFullYear());
  // Card picker month/year for schedule selection
  const [pickerMonth, setPickerMonth] = useState<number>(() => new Date().getMonth() + 1);
  const [pickerYear, setPickerYear] = useState<number>(() => new Date().getFullYear());
  // Fix 1: Add student to past session
  const [showAddStudentPanel, setShowAddStudentPanel] = useState(false);
  const [addStudentId, setAddStudentId] = useState('');
  const [addStudentStatus, setAddStudentStatus] = useState('Present');

  const markedScheduleIds = useMemo(() => {
    return Array.from(new Set(db.studentAttendance.map(a => a.scheduleId)));
  }, [db.studentAttendance]);

  const availableSchedules = useMemo(() => {
     // Sort ASCENDING (earliest date first) for easier tracking
     let scheds = [...(db.calendar || [])].sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
     if (user && user.role === 'tutor') {
        scheds = scheds.filter(c => c.tutor && c.tutor.split(' & ').includes(user.name));
     }
     // HIDE schedules that have already been marked
     return scheds.filter(c => !markedScheduleIds.includes(c.id));
  }, [db.calendar, user, markedScheduleIds]);

  const markedSchedules = useMemo(() => {
     // Sort ASCENDING (earliest date first)
     let scheds = [...(db.calendar || [])].sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
     if (user && user.role === 'tutor') {
        scheds = scheds.filter(c => c.tutor && c.tutor.split(' & ').includes(user.name));
     }
     // SHOW ONLY schedules that have already been marked
     return scheds.filter(c => markedScheduleIds.includes(c.id));
  }, [db.calendar, user, markedScheduleIds]);

  // Filtered by the month/year pre-filter for Edit Records dropdown
  const markedSchedulesFiltered = useMemo(() => {
    const prefix = `${editFilterYear}-${String(editFilterMonth).padStart(2, '0')}`;
    return markedSchedules.filter(c => (c.date || '').startsWith(prefix));
  }, [markedSchedules, editFilterMonth, editFilterYear]);

  // ALL schedules for picker (unmarked + marked), sorted ascending, filtered by month
  const schedulesForPicker = useMemo(() => {
    const prefix = `${pickerYear}-${String(pickerMonth).padStart(2, '0')}`;
    const unmarked = availableSchedules.filter(c => (c.date || '').startsWith(prefix));
    const marked = markedSchedules.filter(c => (c.date || '').startsWith(prefix));
    // Merge and sort ascending by date
    return [...unmarked, ...marked].sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
  }, [availableSchedules, markedSchedules, pickerMonth, pickerYear]);

  const selectedSchedule = useMemo(() => {
     return availableSchedules.find(s => s.id === selectedScheduleId) || null;
  }, [selectedScheduleId, availableSchedules]);

  const activeStudents = useMemo(() => sortStudentsLogically(db.students.filter((s) => s.status === 'Active')), [db.students]);
  
  const targetStudents = useMemo(() => {
     if (!selectedSchedule) return [];
     const sGroup = selectedSchedule.sessionGroup || selectedSchedule.name;
     // Fix 4: Use getStudentSession (which respects sessionOverride) so overridden students
     // are completely excluded from their original session and only appear in the overridden one.
     return activeStudents.filter((s) => getStudentSession(s) === sGroup);
  }, [activeStudents, selectedSchedule]);

  const alreadyMarkedIds = useMemo(() => {
     if (!selectedSchedule) return [];
     const sGroup = selectedSchedule.sessionGroup || selectedSchedule.name;
     return db.studentAttendance
        .filter((a) => a.scheduleId === selectedSchedule.id || (a.date === selectedSchedule.date && a.sessionGroup === sGroup))
        .map((a) => a.studentId);
  }, [db.studentAttendance, selectedSchedule]);

  const studentsToMark = useMemo(() => targetStudents.filter((s) => !alreadyMarkedIds.includes(s.id)), [targetStudents, alreadyMarkedIds]);

  useEffect(() => {
    const initial = {};
    studentsToMark.forEach((s) => (initial[s.id] = 'Present'));
    setAttendanceData(initial);
    // Fix 6: mark dirty when a schedule is selected and students are loaded
    if (setModuleDirty) setModuleDirty(studentsToMark.length > 0 && !!selectedScheduleId);
  }, [studentsToMark]);

  const handleSave = () => {
    if (Object.keys(attendanceData).length === 0 || !selectedSchedule || isSubmitting) return;
    // Fix #7: cegah pengisian absensi retroaktif — hanya izinkan pada hari jadwal berlangsung.
    // Admin dikecualikan agar tetap bisa koreksi data historis.
    const schedDate = selectedSchedule.date;
    const todayLocal = getTodayDateLocal();
    if (user.role === 'tutor' && schedDate !== todayLocal) {
      showToast('Attendance can only be submitted on the scheduled class date', 'warning');
      return;
    }
    setIsSubmitting(true);
    const sGroup = selectedSchedule.sessionGroup || selectedSchedule.name;
    const newRecords = Object.entries(attendanceData).map(([studentId, status]) => {
      const student = activeStudents.find((s) => s.id === studentId);
      return { 
         id: generateId ? generateId('ATT', 'studentAttendance') : `ATT-${Date.now()}-${Math.floor(Math.random() * 1000)}`, 
         scheduleId: selectedSchedule.id,
         date: selectedSchedule.date, 
         sessionGroup: sGroup, 
         studentId, 
         studentName: student.name, 
         class: student.class, 
         status,
         timestamp: getLocalTimestamp(), // FIX BUG #4: diperlukan untuk LWW yang benar di mergeByIds
      };
    });
    // Fix 3: Deduplicate — skip any record for a studentId already marked in this schedule
    setDb((prev) => {
      const existingStudentIdsInSession = new Set(
        prev.studentAttendance
          .filter(a => a.scheduleId === selectedSchedule.id)
          .map(a => a.studentId)
      );
      const deduped = newRecords.filter(r => !existingStudentIdsInSession.has(r.studentId));
      return { ...prev, studentAttendance: [...prev.studentAttendance, ...deduped] };
    });
    showToast('Attendance Saved');
    setAttendanceData({});
    if (setModuleDirty) setModuleDirty(false);
    setTimeout(() => setIsSubmitting(false), 2000);
  };

  const saveEdit = (id) => {
    // BUGFIX (BUG 5): Tambah timestamp agar LWW di mergeByIds tahu versi lokal lebih baru dari cloud
    setDb((p) => ({ ...p, studentAttendance: p.studentAttendance.map((a) => a.id === id ? { ...a, status: editStatus, timestamp: getLocalTimestamp() } : a) }));
    setEditingId(null);
    showToast('Updated');
  };

  const handleDeleteSession = () => {
    if (!editScheduleIdFilter) return;
    const sessionRecords = db.studentAttendance.filter(a => a.scheduleId === editScheduleIdFilter);
    if (sessionRecords.length === 0) return;
    const schedInfo = markedSchedules.find(s => s.id === editScheduleIdFilter);
    const label = schedInfo ? `${schedInfo.date} - ${schedInfo.sessionGroup || schedInfo.name}` : editScheduleIdFilter;
    requestConfirm(
      'Delete Entire Session?',
      `This will move all ${sessionRecords.length} attendance record(s) for "${label}" to the Recycle Bin. This can be undone from the Recycle Bin.`,
      () => {
        const now = getLocalTimestamp();
        const binItems = sessionRecords.map(item => {
          const binId = 'BIN-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
          return { id: binId, binId, originalCollection: 'studentAttendance', deletedAt: now, data: item };
        });
        setDb(prev => ({
          ...prev,
          studentAttendance: prev.studentAttendance.filter(a => a.scheduleId !== editScheduleIdFilter),
          recycleBin: [...(prev.recycleBin || []), ...binItems],
        }));
        setEditScheduleIdFilter('');
        showToast(sessionRecords.length + ' records moved to Recycle Bin', 'warning');
      }
    );
  };

  // Fix 1: Handler to add a student manually to a past session's attendance records
  const handleAddStudentToSession = () => {
    if (!editScheduleIdFilter || !addStudentId) {
      showToast('Please select a past session and a student first', 'warning');
      return;
    }
    const sched = markedSchedules.find(s => s.id === editScheduleIdFilter);
    if (!sched) return;
    const student = db.students.find(s => s.id === addStudentId);
    if (!student) return;
    // Fix 3 inline: prevent duplicate in this session
    const alreadyExists = db.studentAttendance.some(
      a => a.studentId === addStudentId && a.scheduleId === editScheduleIdFilter
    );
    if (alreadyExists) {
      showToast(`${student.name} already has a record in this session`, 'warning');
      return;
    }
    const sGroup = sched.sessionGroup || sched.name;
    const newRecord = {
      id: generateId ? generateId('ATT', 'studentAttendance') : `ATT-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      scheduleId: sched.id,
      date: sched.date,
      sessionGroup: sGroup,
      studentId: addStudentId,
      studentName: student.name,
      class: student.class,
      status: addStudentStatus,
      timestamp: getLocalTimestamp(),
      addedManually: true,
      addedBy: user.name,
    };
    setDb(prev => ({ ...prev, studentAttendance: [...prev.studentAttendance, newRecord] }));
    showToast(`${student.name} added to session with status: ${addStudentStatus}`);
    setAddStudentId('');
    setShowAddStudentPanel(false);
  };

  const visibleAttendanceRecords = useMemo(() => {
    let records = db.studentAttendance;
    if (editScheduleIdFilter) {
        records = records.filter(a => a.scheduleId === editScheduleIdFilter);
    } else {
        records = records.filter((a) => a.date === viewDate);
        if (user && user.role === 'tutor') {
          // Fix #5: hanya tampilkan record dari jadwal yang memang milik tutor ini.
          // Hapus fallback "a.sessionGroup === user.teachingSession" karena bisa bocorkan
          // data absensi sesi lain ke tutor yang kebetulan mengajar sesi yang sama.
          records = records.filter(a => {
              const sched = db.calendar.find(c => c.id === a.scheduleId);
              return sched && sched.tutor && sched.tutor.split(' & ').includes(user.name);
          });
        }
    }
    // Fix 10: Search by student name in Edit Records
    if (editSearchName.trim()) {
      const term = editSearchName.trim().toLowerCase();
      records = records.filter(a => (a.studentName || '').toLowerCase().includes(term));
    }
    return records;
  }, [db.studentAttendance, viewDate, user, editScheduleIdFilter, db.calendar, editSearchName]);

  const statusColors = {
    Present: 'accent-green-500',
    Sick: 'accent-yellow-500',
    Excused: 'accent-purple-500',
    Absent: 'accent-red-500',
  };

  const checkHasDebt = (studentId, dateStr) => {
    const s = db.students.find((x) => x.id === studentId);
    if (!s || s.paymentPlan !== 'Per Visit' || !dateStr) return false;
    const [y, m] = String(dateStr).split('-');
    const monthPrefix = `${y}-${m}`;
    const sPaid = db.payments
      .filter((p) => p.studentId === studentId && Number(p.month) === Number(m) && String(p.year) === String(y) && p.status === 'Paid')
      .reduce((sum, p) => sum + Number(p.amount), 0);
    const studentObj = db.students.find(x => x.id === studentId);
    const sTarget = db.studentAttendance.filter((a) => a.studentId === studentId && a.date.startsWith(monthPrefix) && a.status === 'Present' && isAttendanceValidForStudent(a, studentObj)).length * 25000;
    return sPaid - sTarget < 0;
  };

  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">Student Attendance</h2><p className="text-gray-400 text-sm">Mark and review daily attendance by schedule.</p></div>
      
      <Card className="flex flex-col gap-4 bg-[#0A0E17]">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div>
            <h3 className="font-semibold text-white text-base">Select Schedule to Mark</h3>
            <p className="text-xs text-gray-500 mt-0.5">Choose month, then tap a session card below.</p>
          </div>
          <div className="flex gap-2">
            <select
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none"
              value={pickerMonth}
              onChange={e => { setPickerMonth(Number(e.target.value)); setSelectedScheduleId(''); }}
            >
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input
              type="number"
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none w-24"
              value={pickerYear}
              onChange={e => { setPickerYear(Number(e.target.value)); setSelectedScheduleId(''); }}
            />
          </div>
        </div>

        {availableSchedules.length === 0 && markedSchedules.length === 0 ? (
          <div className="w-full p-4 bg-red-500/10 border border-red-500/20 rounded-lg text-red-400 font-medium text-sm">
            No Academic Calendar schedule available. Please create a schedule first.
          </div>
        ) : schedulesForPicker.length === 0 ? (
          <div className="w-full p-4 bg-amber-500/10 border border-amber-500/20 rounded-lg text-amber-400 font-medium text-sm">
            No sessions in {MONTHS[pickerMonth - 1]} {pickerYear}. Try a different month.
          </div>
        ) : (
          <div className="max-h-[420px] overflow-y-auto custom-scrollbar pr-1 -mr-1">
            <div className="grid grid-cols-2 sm:grid-cols-2 lg:grid-cols-3 gap-2 sm:gap-3">
              {schedulesForPicker.map(c => {
                const isMarkedDone = markedScheduleIds.includes(c.id);
                const isSelected = selectedScheduleId === c.id;
                const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
                const [y, mo, d] = (c.date || '').split('-');
                const dayName = c.date ? dayNames[new Date(Number(y), Number(mo)-1, Number(d)).getDay()] : '';
                return (
                  <button
                    key={c.id}
                    onClick={() => !isMarkedDone && setSelectedScheduleId(isSelected ? '' : c.id)}
                    disabled={isMarkedDone}
                    className={`text-left p-2.5 sm:p-4 rounded-xl border-2 transition-all duration-150 relative ${
                      isMarkedDone
                        ? 'border-emerald-700/60 bg-emerald-950/40 cursor-default opacity-80'
                        : isSelected
                          ? 'border-[#00D4FF] bg-[#00D4FF]/10 shadow-[0_0_20px_rgba(0,212,255,0.2)]'
                          : 'border-gray-700 bg-[#151B26] hover:border-gray-500 hover:bg-[#1A2234]'
                    }`}
                  >
                    {isMarkedDone && (
                      <span className="absolute top-2 right-2 sm:top-3 sm:right-3 flex items-center gap-1 bg-emerald-500/20 border border-emerald-500/40 text-emerald-400 text-[9px] font-bold px-1.5 py-0.5 rounded-full">
                        <CheckCircle2 size={9} />
                        Done
                      </span>
                    )}
                    <div className="flex items-start justify-between gap-1 mb-1.5">
                      <div className="min-w-0">
                        <p className={`text-xs sm:text-sm font-bold leading-tight ${isMarkedDone ? 'text-emerald-400' : isSelected ? 'text-[#00D4FF]' : 'text-white'}`}>
                          <span className="hidden sm:inline">{dayName}, </span>{parseInt(d, 10)} {MONTHS[parseInt(mo, 10) - 1]} {y}
                        </p>
                        <p className="text-[10px] text-gray-500 mt-0.5 sm:hidden">{dayName}</p>
                      </div>
                      {isSelected && !isMarkedDone && (
                        <span className="flex-shrink-0 w-4 h-4 sm:w-5 sm:h-5 rounded-full bg-[#00D4FF] flex items-center justify-center">
                          <Check size={10} className="text-[#0B0F19]" strokeWidth={3} />
                        </span>
                      )}
                    </div>
                    <p className={`text-[10px] sm:text-xs font-semibold px-1.5 py-0.5 rounded-full inline-block mb-1 truncate max-w-full ${
                      isMarkedDone
                        ? 'text-emerald-400/80 bg-emerald-500/10'
                        : 'text-[#00D4FF]/80 bg-[#00D4FF]/10'
                    }`}>
                      {c.sessionGroup || c.name}
                    </p>
                    <div className="flex flex-col sm:flex-row sm:items-center gap-0.5 sm:gap-3 text-[10px] sm:text-xs text-gray-400 mt-0.5">
                      <span className="flex items-center gap-1 truncate"><User size={10} /><span className="truncate">{c.tutor}</span></span>
                      {c.startTime && <span className="flex items-center gap-1 flex-shrink-0"><Clock size={10} />{c.startTime}{c.endTime ? ` - ${c.endTime}` : ''}</span>}
                    </div>
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </Card>

      <Card className="p-0 overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr>
              <th className="p-4 text-center w-12 text-gray-400">No.</th>
              <th className="p-4">Name</th>
              <th className="p-4 text-center">Present</th>
              <th className="p-4 text-center">Sick</th>
              <th className="p-4 text-center">Excused</th>
              <th className="p-4 text-center">Absent</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {studentsToMark.length > 0 ? studentsToMark.map((s, index) => (
              <tr key={s.id} className="hover:bg-[#0B0F19]">
                <td className="p-4 text-center"><span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF]">{String(index + 1).padStart(2, '0')}</span></td>
                <td className="p-4 text-white font-medium"><div className="flex items-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /> <span className="text-xs text-gray-500 ml-1.5">({s.class})</span>{checkHasDebt(s.id, selectedSchedule?.date) && <span title="Memiliki tunggakan pembayaran"><AlertCircle size={14} className="text-amber-500 ml-2" /></span>}</div></td>
                {['Present', 'Sick', 'Excused', 'Absent'].map((status) => (
                  <td key={status} className="p-4 text-center">
                    <input type="radio" checked={attendanceData[s.id] === status} onChange={() => setAttendanceData((p) => ({ ...p, [s.id]: status }))} className={`w-5 h-5 cursor-pointer ${statusColors[status]}`} />
                  </td>
                ))}
              </tr>
            )) : (
              <tr>
                <td colSpan={6}>
                  <EmptyState
                    icon={!selectedSchedule ? CalendarIcon : CheckCircle2}
                    title={!selectedSchedule ? "No schedule selected" : "All students marked"}
                    description={!selectedSchedule ? "Please select a schedule to mark attendance." : "All active students for this schedule have been marked."}
                  />
                </td>
              </tr>
            )}
          </tbody>
        </table>
        <div className="p-4 bg-[#0A0E17] flex justify-end border-t border-gray-800">
          <Button
            onClick={handleSave}
            disabled={studentsToMark.length === 0 || !selectedSchedule || isSubmitting}
            icon={isSubmitting ? RefreshCw : CheckCircle2}
            className={isSubmitting ? 'opacity-80' : ''}
          >
            {isSubmitting ? 'Saving...' : 'Save Attendance'}
          </Button>
        </div>
      </Card>
      
      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col gap-3">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
            <h3 className="font-semibold text-white">Edit Records</h3>
            {/* Action buttons shown only when a session is selected */}
            {editScheduleIdFilter && (
              <div className="flex gap-2 flex-wrap">
                <button
                  onClick={() => setShowAddStudentPanel(p => !p)}
                  className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[#00D4FF]/10 border border-[#00D4FF]/30 text-[#00D4FF] hover:bg-[#00D4FF]/20 transition-all text-sm font-semibold whitespace-nowrap"
                >
                  <Plus size={14} /> Add Student
                </button>
                <button
                  onClick={handleDeleteSession}
                  className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-red-500/10 border border-red-500/30 text-red-400 hover:bg-red-500/20 hover:border-red-500/60 transition-all text-sm font-semibold whitespace-nowrap"
                >
                  <Trash2 size={14} /> Delete Session
                </button>
              </div>
            )}
          </div>

          {/* Step 1: Pick Month & Year to narrow the list */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-gray-500 font-medium uppercase tracking-wider whitespace-nowrap">Step 1 — Month:</span>
            <select
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none"
              value={editFilterMonth}
              onChange={e => { setEditFilterMonth(Number(e.target.value)); setEditScheduleIdFilter(''); setShowAddStudentPanel(false); setAddStudentId(''); }}
            >
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input
              type="number"
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none w-24"
              value={editFilterYear}
              onChange={e => { setEditFilterYear(Number(e.target.value)); setEditScheduleIdFilter(''); setShowAddStudentPanel(false); setAddStudentId(''); }}
            />
            <span className="text-xs text-gray-600">
              {markedSchedulesFiltered.length === 0 ? 'No sessions recorded this month' : `${markedSchedulesFiltered.length} session(s) found`}
            </span>
          </div>

          {/* Step 2: Pick specific session from the short list OR view by date */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-gray-500 font-medium uppercase tracking-wider whitespace-nowrap">Step 2 — Session:</span>
            <select
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none flex-1 min-w-[220px]"
              value={editScheduleIdFilter}
              onChange={e => { setEditScheduleIdFilter(e.target.value); setShowAddStudentPanel(false); setAddStudentId(''); setEditSearchName(''); }}
              disabled={markedSchedulesFiltered.length === 0}
            >
              <option value="">— Select session —</option>
              {markedSchedulesFiltered.map(c => (
                <option key={c.id} value={c.id}>
                  {c.date} • {c.sessionGroup || c.name}
                </option>
              ))}
            </select>
            {!editScheduleIdFilter && (
              <>
                <span className="text-xs text-gray-600">or view by date:</span>
                <input type="date" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none" value={viewDate} onChange={e => setViewDate(e.target.value)} />
              </>
            )}
          </div>

          {/* Fix 10: Search by student name in Edit Records */}
          <div className="flex items-center gap-2">
            <span className="text-xs text-gray-500 font-medium uppercase tracking-wider whitespace-nowrap">Step 3 — Search:</span>
            <div className="relative flex-1 max-w-xs">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none" />
              <input
                type="text"
                placeholder="Filter by student name..."
                value={editSearchName}
                onChange={e => setEditSearchName(e.target.value)}
                className="w-full bg-[#151B26] border border-gray-700 rounded-lg pl-8 pr-3 py-1.5 text-white text-sm focus:border-[#00D4FF] focus:outline-none placeholder-gray-600"
              />
              {editSearchName && (
                <button onClick={() => setEditSearchName('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-500 hover:text-white transition-colors">
                  <X size={14} />
                </button>
              )}
            </div>
            {editSearchName && <span className="text-xs text-gray-500">{visibleAttendanceRecords.length} found</span>}
          </div>
        </div>
        {/* Fix 1: Add Student to Session Panel */}
        {showAddStudentPanel && editScheduleIdFilter && (
          <div className="p-4 bg-[#0A0E17] border-b border-[#00D4FF]/30 flex flex-col sm:flex-row gap-3 items-end">
            <div className="flex-1">
              <label className="block text-xs text-gray-400 mb-1.5 font-medium">Select Registered Student to Add</label>
              <select
                className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:border-[#00D4FF] focus:outline-none"
                value={addStudentId}
                onChange={e => setAddStudentId(e.target.value)}
              >
                <option value="">Choose a student...</option>
                {sortStudentsLogically(db.students.filter(s => s.status === 'Active')).map(s => (
                  <option key={s.id} value={s.id}>{s.name} ({s.class})</option>
                ))}
              </select>
            </div>
            <div className="w-full sm:w-40">
              <label className="block text-xs text-gray-400 mb-1.5 font-medium">Status</label>
              <select
                className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:border-[#00D4FF] focus:outline-none"
                value={addStudentStatus}
                onChange={e => setAddStudentStatus(e.target.value)}
              >
                <option>Present</option><option>Sick</option><option>Excused</option><option>Absent</option>
              </select>
            </div>
            <div className="flex gap-2">
              <Button onClick={handleAddStudentToSession} icon={Plus} disabled={!addStudentId}>Add</Button>
              <Button variant="secondary" onClick={() => { setShowAddStudentPanel(false); setAddStudentId(''); }}>Cancel</Button>
            </div>
          </div>
        )}
        <div className="p-4 space-y-2 bg-[#151B26]">
          {visibleAttendanceRecords.map((r) => (
            <div key={r.id} className="flex justify-between items-center p-3 bg-[#0B0F19] rounded-lg border border-gray-800">
              <div>
                <p className="text-white text-sm flex items-center">{r.studentName}{checkHasDebt(r.studentId, r.date) && <span title="Memiliki tunggakan pembayaran"><AlertCircle size={14} className="text-amber-500 ml-2" /></span>}</p>
                <p className="text-xs text-gray-500">{r.class} • {r.sessionGroup}</p>
              </div>
              <div className="flex gap-2 items-center text-center">
                {editingId === r.id ? (
                  <>
                    <select className="bg-[#151B26] border border-gray-700 rounded text-sm text-white p-1" value={editStatus} onChange={(e) => setEditStatus(e.target.value)}>
                      <option>Present</option><option>Sick</option><option>Excused</option><option>Absent</option>
                    </select>
                    <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => saveEdit(r.id)}>Save</Button>
                  </>
                ) : (
                  <>
                    <Badge status={r.status} />
                    <button onClick={() => { setEditingId(r.id); setEditStatus(r.status); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Record"><Edit2 size={18} /></button>
                    <button onClick={() => softDelete('studentAttendance', r.id, `Attendance for ${r.studentName}`) } className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Record"><Trash2 size={18} /></button>
                  </>
                )}
              </div>
            </div>
          ))}
          {visibleAttendanceRecords.length === 0 && (
             <EmptyState icon={UserCheck} title="No attendance records found" description="Try adjusting the date or session filter." className="py-6" />
          )}
        </div>
      </Card>
    </div>
  );
}

function TutorAttendanceModule({ db, setDb, user, showToast, softDelete, generateId }) {
  const [editingId, setEditingId] = useState(null);
  const [editStatus, setEditStatus] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Fix 8: Persistent month/year filters via localStorage
  const [filterMonth, setFilterMonth] = useLocalStorage<number | string>('ecg_tutoratt_month', new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useLocalStorage<number>('ecg_tutoratt_year', new Date().getFullYear());
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [filterMonth, filterYear, rowsPerPage]);

  const handleCheckIn = () => {
    if (isSubmitting) return;
    const dateObj = new Date();
    // Gunakan waktu lokal agar tidak terjadi lompatan hari zona waktu (UTC issue)
    const today = getTodayDateLocal();
    const day = dateObj.toLocaleDateString('en-US', { weekday: 'long' });
    const time = dateObj.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

    const existingRecords = db.tutorAttendance || [];
    // BUG FIX #4: Also match legacy records that only store name (no tutorId) to prevent duplicates
    if (existingRecords.find((a) => (a.tutorId === user.id || (!a.tutorId && a.name === user.name)) && a.date === today)) {
       return showToast('Already checked in today', 'warning');
    }
    // Fix #2: only allow check-in when there is a scheduled class today
    const hasScheduleToday = (db.calendar || []).some(c => c.date === today && c.tutor && c.tutor.split(' & ').includes(user.name));
    if (!hasScheduleToday) {
       return showToast('No scheduled class today — check-in not allowed', 'warning');
    }
    
    setIsSubmitting(true);
    setDb((prev) => ({ 
      ...prev, 
      tutorAttendance: [
        ...(prev.tutorAttendance || []), 
        { id: generateId ? generateId('TA', 'tutorAttendance') : `TA-${Date.now()}-${Math.random().toString(36).slice(2,6).toUpperCase()}`, tutorId: user.id, name: user.name, date: today, day, time, status: 'Present' }
      ] 
    }));
    showToast('Checked in successfully');
    setTimeout(() => setIsSubmitting(false), 2000);
  };

  const saveEdit = (id) => {
    setDb((p) => ({
      ...p,
      tutorAttendance: p.tutorAttendance.map((a) =>
        a.id === id
          ? {
              ...a,
              status: editStatus,
              // Fix #3: audit trail — record who changed the status and when
              lastEditedBy: user.name,
              lastEditedAt: getLocalTimestamp(),
              originalStatus: a.originalStatus || a.status,
            }
          : a
      ),
    }));
    setEditingId(null);
    showToast('Updated');
  };

  const tutorRecords = db.tutorAttendance || [];
  
  const filteredRecords = useMemo(() => {
    let records = tutorRecords.filter(a => {
       if (filterMonth !== 'All') {
          const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
          return a.date.startsWith(prefix);
       }
       return a.date.startsWith(String(filterYear));
    });
    if (user.role !== 'admin') {
       records = records.filter(a => a.tutorId === user.id);
    }
    return records.slice().reverse();
  }, [tutorRecords, filterMonth, filterYear, user]);

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredRecords.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredRecords.length / (rowsNum || 1));
  const paginatedData = isAll ? filteredRecords : filteredRecords.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  return (
    <div className="space-y-6">
      {/* Fix 3: Standardized module header */}
      <div>
        <h2 className="text-2xl font-bold text-white mb-1">Tutor Attendance</h2>
        <p className="text-gray-400 text-sm">Record and review tutor check-in history by month.</p>
      </div>

      <Card className="text-center py-12 border border-[#00D4FF]/15">
        <div className="w-16 h-16 rounded-2xl bg-[#00D4FF]/10 border border-[#00D4FF]/20 flex items-center justify-center mx-auto mb-4">
          <Activity size={28} className="text-[#00D4FF]" />
        </div>
        <h2 className="text-xl font-bold text-white mb-1">Tutor Self Check-In</h2>
        <p className="text-gray-500 text-sm mb-5">Tap the button below to record your attendance for today's scheduled class.</p>
        <Button
          onClick={handleCheckIn}
          disabled={isSubmitting}
          className="mx-auto py-3 px-8 text-lg"
          icon={isSubmitting ? RefreshCw : CheckCircle2}
        >
          {isSubmitting ? 'Recording...' : 'Check In Now'}
        </Button>
      </Card>

      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-3 flex-wrap">
           <h3 className="font-semibold text-white w-full sm:w-auto">{user.role === 'admin' ? 'All Tutors History' : 'My Check-In History'}</h3>
           <div className="flex gap-2 w-full sm:w-auto">
             <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
               <option value="All">All Months</option>
               {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
             </select>
             <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
           </div>
        </div>
        
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
              <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">Date</th><th className="p-4 text-center">Day</th><th className="p-4 text-center">Time</th>{user.role === 'admin' && <th className="p-4 text-center">Tutor</th>}<th className="p-4 text-center">Status</th><th className="p-4 text-center">Actions</th></tr>
            </thead>
            <tbody className="divide-y divide-gray-800">
              {paginatedData.map((a, index) => (
                <tr key={a.id} className="hover:bg-[#0B0F19]">
                  <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                  <td className="p-4 text-center">{a.date}</td>
                  <td className="p-4 text-center">{a.day}</td>
                  <td className="p-4 text-center">{a.time}</td>
                  {user.role === 'admin' && <td className="p-4 text-center">{a.name}</td>}
                  <td className="p-4 text-center">
                    {editingId === a.id ? (
                      <select className="bg-[#151B26] border border-gray-700 rounded text-sm text-white p-1" value={editStatus} onChange={(e) => setEditStatus(e.target.value)}>
                        <option>Present</option><option>Sick</option><option>Excused</option><option>Absent</option>
                      </select>
                    ) : (
                      <Badge status={a.status} />
                    )}
                  </td>
                  <td className="p-4 text-center flex justify-center gap-2">
                    {editingId === a.id ? (
                      <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => saveEdit(a.id)}>Save</Button>
                    ) : (
                      <>
                        <button onClick={() => { setEditingId(a.id); setEditStatus(a.status); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Record"><Edit2 size={18} /></button>
                        <button onClick={() => softDelete('tutorAttendance', a.id, 'Tutor Attendance') } className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Record"><Trash2 size={18} /></button>
                      </>
                    )}
                  </td>
                </tr>
              ))}
              {paginatedData.length === 0 && (
               <tr><td colSpan={7}><EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." /></td></tr>
            )}
          </tbody>
        </table>
        </div>

        {/* Pagination Footer */}
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
           <div className="flex items-center gap-2">
             <span>Show</span>
             <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
               <option value={10}>10</option>
               <option value={20}>20</option>
               <option value={50}>50</option>
               <option value="All">All</option>
             </select>
             <span>entries {filteredRecords.length > 0 && `(Total: ${filteredRecords.length})`}</span>
           </div>
           
           {!isAll && totalPages > 1 && (
             <div className="flex items-center gap-2">
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
               <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
             </div>
           )}
        </div>
      </Card>
    </div>
  );
}

function AssessmentsModule({ db, setDb, generateId, showToast, user }) {
  const defaultSession = user?.role === 'tutor' ? (parseSessions(user.teachingSession)[0] || SESSIONS[0]) : SESSIONS[0];
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [year, setYear] = useState(new Date().getFullYear());
  const [sessionGroup, setSessionGroup] = useState(defaultSession);
  const [tableData, setTableData] = useState({});
  const [materialFilterSession, setMaterialFilterSession] = useState('All');

  const activeStudents = useMemo(() => sortStudentsLogically(db.students.filter((s) => s.status === 'Active')), [db.students]);
  const targetStudents = useMemo(() => activeStudents.filter((s) => getStudentSession(s) === sessionGroup), [activeStudents, sessionGroup]);

  const subjects = useMemo(() => {
    return sessionGroup === SESSIONS[0]
      ? ['Reading', 'Writing', 'Math', 'English']
      : ['Speaking', 'Writing', 'Reading', 'Listening'];
  }, [sessionGroup]);

  const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
  const allMeetingsThisMonth = useMemo(() => {
    return db.journals.filter(j => (j.date || '').startsWith(monthPrefix)).sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
  }, [db.journals, monthPrefix]);

  const filteredMeetings = useMemo(() => {
    if (materialFilterSession === 'All') return allMeetingsThisMonth;
    return allMeetingsThisMonth.filter(j => j.sessionGroup === materialFilterSession);
  }, [allMeetingsThisMonth, materialFilterSession]);

  useEffect(() => {
    const initialData = {};
    targetStudents.forEach((student) => {
      const existing = db.assessments.find((a) => a.studentId === student.id && Number(a.month) === Number(month) && Number(a.year) === Number(year));
      if (existing && existing.scores) {
        initialData[student.id] = { ...existing.scores };
      } else {
        initialData[student.id] = {};
        subjects.forEach((sub) => (initialData[student.id][sub] = ''));
      }
    });
    setTableData(initialData);
  }, [month, year, sessionGroup, targetStudents, db.assessments, subjects]);

  const handleScoreChange = (studentId, field, val) => {
    let num = val === '' ? '' : parseInt(val, 10);
    if (num !== '' && Number.isNaN(num as number)) return;
    if (num > 100) num = 100;
    if (num < 0) num = 0;
    setTableData((prev) => ({ ...prev, [studentId]: { ...prev[studentId], [field]: num } }));
  };

  const handleSaveAll = () => {
    // Fix #8: lock periode — tutor tidak boleh mengubah nilai bulan yang sudah lewat.
    // Admin tetap bisa edit kapan saja untuk keperluan koreksi.
    if (user?.role === 'tutor') {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      if (Number(year) < currentYear || (Number(year) === currentYear && Number(month) < currentMonth)) {
        showToast('Cannot edit assessments for past months', 'warning');
        return;
      }
    }
    let newAssessments = [...db.assessments];
    let updatedCount = 0;

    targetStudents.forEach((student) => {
      const studentScores = tableData[student.id];
      if (!studentScores) return;

      const hasScores = subjects.some((sub) => studentScores[sub] !== '' && studentScores[sub] !== undefined);
      if (!hasScores) return;

      let total = 0, count = 0;
      subjects.forEach((sub) => {
        if (studentScores[sub] !== '' && studentScores[sub] !== undefined) {
          total += Number(studentScores[sub]);
          count++;
        }
      });
      const average = count > 0 ? Math.round(total / count) : 0;
      const grade = average >= 90 ? 'A' : average >= 80 ? 'B' : average >= 70 ? 'C' : 'D';

      const existingIdx = newAssessments.findIndex((a) => a.studentId === student.id && Number(a.month) === Number(month) && Number(a.year) === Number(year));

      const cleanedScores = {};
      subjects.forEach((sub) => cleanedScores[sub] = studentScores[sub] === '' ? '' : Number(studentScores[sub]));

      const assessmentRecord = {
        id: existingIdx >= 0 ? newAssessments[existingIdx].id : generateId('ASS', 'assessments'),
        studentId: student.id, studentName: student.name, level: student.level, class: student.class, month: String(month), year: String(year), sessionGroup: sessionGroup, scores: cleanedScores, average, grade,
        updatedAt: getSyncTimestamp(), // FIX BUG 3: LWW butuh timestamp agar mergeByIds bisa memilih versi terbaru
      };

      if (existingIdx >= 0) newAssessments[existingIdx] = assessmentRecord;
      else newAssessments.push(assessmentRecord);
      updatedCount++;
    });

    if (updatedCount > 0) {
      setDb((prev) => ({ ...prev, assessments: newAssessments }));
      showToast(`Assessment saved successfully for ${updatedCount} students.`);
    } else {
      showToast('No scores entered to save.', 'warning');
    }
  };

  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">Monthly Assessment</h2><p className="text-gray-400 text-sm">Spreadsheet-style mass grading. Select a period to grade or edit existing records.</p></div>
      <div className="flex flex-col md:flex-row gap-3 items-stretch md:items-center bg-[#0A0E17] p-4 sm:p-5 rounded-xl border border-gray-800 shadow-sm flex-wrap">
        <Input label="Month" type="select" options={MONTHS.map((m, i) => ({ value: i + 1, label: m }))} value={month} onChange={setMonth} className="mb-0" />
        <Input label="Year" type="number" value={year} onChange={setYear} className="mb-0" />
        <Input label="Session" type="select" options={user?.role === 'tutor' ? parseSessions(user.teachingSession) : SESSIONS} value={sessionGroup} onChange={setSessionGroup} className="mb-0" />
      </div>

      <div className="bg-[#151B26] border border-gray-800 rounded-xl shadow-sm overflow-hidden mb-6">
        <div className="p-4 border-b border-gray-800 bg-[#0A0E17] flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <h3 className="text-white font-bold text-sm">Learning Materials (From Journals)</h3>
          <select
            className="bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-1.5 text-white text-sm focus:outline-none focus:border-[#00D4FF]"
            value={materialFilterSession}
            onChange={e => setMaterialFilterSession(e.target.value)}
          >
            {user?.role === 'tutor' ? (
               <>
                 <option value="All">All My Sessions</option>
                 {parseSessions(user.teachingSession).map(s => <option key={s} value={s}>{s}</option>)}
               </>
            ) : (
               <>
                 <option value="All">All Sessions</option>
                 {SESSIONS.map(s => <option key={s} value={s}>{s}</option>)}
               </>
            )}
          </select>
        </div>
        <div className="p-4 max-h-48 overflow-y-auto custom-scrollbar bg-[#0B0F19]">
          {filteredMeetings.length > 0 ? (
            <ul className="space-y-3">
              {filteredMeetings.map((m) => (
                <li key={m.id} className="text-sm bg-[#151B26] p-3 rounded-lg border border-gray-800">
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-[#00D4FF] font-medium">{m.sessionGroup}</span>
                    {/* FIX #12: Gunakan safeDateDisplay agar tidak ada UTC-1-day bug */}
                    <span className="text-gray-500 text-xs">{safeDateDisplay(m.date, 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })}</span>
                  </div>
                  <span className="text-gray-300 font-medium">Topic: </span><span className="text-white">{m.topic}</span>
                </li>
              ))}
            </ul>
          ) : (
            <EmptyState icon={BookOpen} title="No materials recorded" description="No Learning Journal entries for this month." className="py-6" />
          )}
        </div>
      </div>

      <div className="bg-[#151B26] border border-gray-800 rounded-xl shadow-xl overflow-hidden">
        <div className="p-4 border-b border-gray-800 bg-[#0A0E17] flex justify-between items-center"><h3 className="text-white font-bold">Grading Table: {sessionGroup}</h3><span className="text-gray-400 text-sm font-medium">{MONTHS[month - 1]} {year}</span></div>

        <div className="overflow-x-auto custom-scrollbar">
          <table className="w-full text-left border-collapse min-w-max">
            <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold sticky top-0 z-10">
              <tr>
                <th className="py-3 px-4 text-center w-12 text-gray-400 font-medium text-sm">No.</th>
                <th className="py-3 px-4 text-gray-400 font-medium text-sm whitespace-nowrap">Student Name</th>
                <th className="py-3 px-4 text-gray-400 font-medium text-sm whitespace-nowrap">Class</th>
                {subjects.map((s) => <th key={s} className="py-3 px-4 text-gray-400 font-medium text-sm text-center w-[100px]">{s}</th>)}
                <th className="py-3 px-4 text-gray-400 font-medium text-sm text-center w-[80px]">Avg</th>
                <th className="py-3 px-4 text-gray-400 font-medium text-sm text-center w-[80px]">Grade</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-800 bg-[#151B26]">
              {targetStudents.map((student, index) => {
                const studentScores = tableData[student.id] || {};
                let total = 0, count = 0;
                subjects.forEach((sub) => { if (studentScores[sub] !== '' && studentScores[sub] !== undefined) { total += Number(studentScores[sub]); count++; } });
                const avg = count > 0 ? Math.round(total / count) : '-';
                const grade = avg === '-' ? '-' : avg >= 90 ? 'A' : avg >= 80 ? 'B' : avg >= 70 ? 'C' : 'D';

                return (
                  <tr key={student.id} className="hover:bg-[#0B0F19] transition-colors" style={{ height: '50px' }}>
                    <td className="py-1 px-4 text-center text-gray-500 font-medium">{index + 1}</td>
                    <td className="py-1 px-4 text-white font-medium whitespace-nowrap"><div className="flex items-center">{student.name} <NewBadge isNew={student.enrollmentStatus} billingStartMonth={student.billingStartMonth || null} /></div></td>
                    <td className="py-1 px-4 text-gray-400 text-xs whitespace-nowrap">{student.class}</td>
                    {subjects.map((sub) => (
                      <td key={sub} className="py-1 px-4 text-center">
                        <input type="number" min="0" max="100" placeholder="0-100" className="w-[80px] h-8 text-center bg-[#0B0F19] border border-gray-700 rounded text-white text-sm focus:outline-none focus:border-[#00D4FF] focus:ring-1 focus:ring-[#00D4FF] transition-all" value={studentScores[sub] !== undefined ? studentScores[sub] : ''} onChange={(e) => handleScoreChange(student.id, sub, e.target.value)} />
                      </td>
                    ))}
                    <td className="py-1 px-4 text-center font-bold text-[#00D4FF]">{avg}</td>
                    <td className="py-1 px-4 text-center font-bold text-gray-300">{grade}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800">
          <Button onClick={handleSaveAll} className="w-full py-4 text-lg font-bold shadow-[0_0_20px_rgba(0,212,255,0.2)]" disabled={targetStudents.length === 0}>SAVE ALL ASSESSMENTS</Button>
        </div>
      </div>
    </div>
  );
}

function PaymentsModule({ db, setDb, generateId, showToast, handlePrint, handleShareImage, downloadPNG, softDelete, language = 'en', isDbDirty }) {
  // Fix 8: Persistent month/year filter
  const [month, setMonth] = useLocalStorage<number>('ecg_payments_month', new Date().getMonth() + 1);
  const [year, setYear] = useLocalStorage<number>('ecg_payments_year', new Date().getFullYear());
  const [sessionGroup, setSessionGroup] = useState('All Sessions');
  
  // NEW: State for Filters & Pagination
  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearchTerm = useDebounce(searchTerm, 300);
  const [filterLevel, setFilterLevel] = useState('');
  const [filterClass, setFilterClass] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  const [selectedInvoice, setSelectedInvoice] = useState(null);
  const [viewingStudentId, setViewingStudentId] = useState(null);
  const [amounts, setAmounts] = useState({});
  const [methods, setMethods] = useState({});
  const [showQuickOverview, setShowQuickOverview] = useState(false);
  const [overviewFilter, setOverviewFilter] = useState<'all' | 'paid' | 'unpaid' | 'partial'>('all');

  useEffect(() => {
    setCurrentPage(1);
  }, [debouncedSearchTerm, filterLevel, filterClass, sessionGroup, month, year, rowsPerPage]);

  const activeStudents = useMemo(() => sortStudentsLogically(db.students.filter((s) => s.status === 'Active')), [db.students]);

  // Hitung kehadiran bulan yang sedang dilihat per siswa (Present saja)
  // Fix 4: Only count attendance records that belong to the student's current effective session
  const attendanceThisMonth = useMemo(() => {
    const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
    const map: Record<string, number> = {};
    (db.studentAttendance || []).forEach(a => {
      if (String(a.date || '').startsWith(monthPrefix) && a.status === 'Present') {
        const studentObj = db.students.find(s => s.id === a.studentId);
        if (isAttendanceValidForStudent(a, studentObj)) {
          map[a.studentId] = (map[a.studentId] || 0) + 1;
        }
      }
    });
    return map;
  }, [db.studentAttendance, db.students, month, year]);
  
  const filteredStudents = useMemo(() => {
    return activeStudents.filter((s) => {
      const matchSearch = (s.name || '').toLowerCase().includes((debouncedSearchTerm || '').toLowerCase());
      const matchSession = sessionGroup === 'All Sessions' ? true : getStudentSession(s) === sessionGroup;
      const matchLevel = filterLevel ? s.level === filterLevel : true;
      const matchClass = filterClass ? s.class === filterClass : true;
      return matchSearch && matchSession && matchLevel && matchClass;
    });
  // FIX #6: Dependency harus debouncedSearchTerm (bukan searchTerm mentah) agar debounce bekerja
  }, [activeStudents, debouncedSearchTerm, sessionGroup, filterLevel, filterClass]);

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredStudents.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredStudents.length / (rowsNum || 1));
  const paginatedData = isAll ? filteredStudents : filteredStudents.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  // Calculate Revenue based on filtered list so Admin can see revenue per specific class/level
  const totalRevenue = useMemo(() => {
    const validStudentIds = new Set(filteredStudents.map(s => s.id));
    return db.payments
      .filter(p => Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid' && validStudentIds.has(p.studentId))
      .reduce((sum, p) => sum + Number(p.amount), 0);
  }, [db.payments, month, year, filteredStudents]);

  // Calculate total outstanding (unpaid + partial deficit) for the selected month
  const revenueStats = useMemo(() => {
    let totalOutstanding = 0;
    let unpaidCount = 0;
    let partialCount = 0;
    let paidCount = 0;
    let totalBilled = 0;
    filteredStudents.forEach(s => {
      const target = (() => {
        if (s.paymentPlan === 'Free') return 0;
        if (s.billingStartMonth) {
          const [bY, bM] = s.billingStartMonth.split('-').map(Number);
          if (bY && bM && (Number(year) < bY || (Number(year) === bY && Number(month) < bM))) return 0;
        }
        const sGroup = getStudentSession(s);
        const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
        if ((s.paymentPlan || 'Monthly') === 'Monthly') {
          return db.calendar.filter(c => c.date.startsWith(monthPrefix) && (c.sessionGroup || c.name) === sGroup).length * 25000;
        }
        return db.studentAttendance.filter(a => a.studentId === s.id && a.date.startsWith(monthPrefix) && a.status === 'Present' && isAttendanceValidForStudent(a, s)).length * 25000;
      })();
      if (target === 0) return;
      const paid = db.payments
        .filter(p => String(p.studentId).trim().toLowerCase() === String(s.id).trim().toLowerCase() && Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid')
        .reduce((sum, p) => sum + Number(p.amount), 0);
      totalBilled += target;
      const deficit = target - paid;
      if (deficit > 0) totalOutstanding += deficit;
      if (paid >= target) paidCount++;
      else if (paid > 0) partialCount++;
      else unpaidCount++;
    });
    const collectionRate = totalBilled > 0 ? Math.round((totalRevenue / totalBilled) * 100) : 0;
    return { totalOutstanding, unpaidCount, partialCount, paidCount, totalBilled, collectionRate };
  }, [filteredStudents, db.payments, db.calendar, db.studentAttendance, month, year, totalRevenue]);

  const getStudentTarget = (student) => {
    // BILLING START CHECK: Jika billingStartMonth diset dan bulan/tahun yang sedang dilihat
    // masih SEBELUM billingStartMonth, siswa ini belum dihitung dalam expected collected.
    if (student.billingStartMonth) {
      const [bYear, bMonth] = student.billingStartMonth.split('-').map(Number);
      if (bYear && bMonth) {
        const viewYear = Number(year);
        const viewMonth = Number(month);
        if (viewYear < bYear || (viewYear === bYear && viewMonth < bMonth)) {
          return 0; // Belum mulai tagih — tidak masuk hitungan
        }
      }
    }
    const plan = student.paymentPlan || 'Monthly';
    // FREE plan: target selalu 0, tidak pernah ditagih
    if (plan === 'Free') return 0;
    const sGroup = getStudentSession(student);
    const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
    if (plan === 'Monthly') {
        return db.calendar.filter(c => c.date.startsWith(monthPrefix) && (c.sessionGroup || c.name) === sGroup).length * 25000;
    } else {
        // Fix 4: Only count attendance that belongs to student's current (effective) session
        return db.studentAttendance.filter(a => a.studentId === student.id && a.date.startsWith(monthPrefix) && a.status === 'Present' && isAttendanceValidForStudent(a, student)).length * 25000;
    }
  };

  const handleRecordInline = (student, amountOverride) => {
    const amt = amountOverride || amounts[student.id];
    if (!amt || amt <= 0) return showToast('Enter valid amount', 'error');
    const sSession = getStudentSession(student);
    const method = methods[student.id] || 'Cash';
    const now = new Date();
    const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const rec = {
      id: generateId('INV', 'payments'),
      studentId: student.id,
      studentName: student.name,
      level: student.level,
      class: student.class,
      sessionGroup: sSession,
      amount: Number(amt), // BUGFIX: Pastikan amount tersimpan sebagai Number, bukan string
      month: String(month),
      year: String(year),
      date: getTodayDateLocal(),
      time: timeStr,
      method: method,
      status: 'Paid',
      timestamp: getLocalTimestamp(), // FIX BUG #3: diperlukan untuk LWW yang benar di mergeByIds
      updatedAt: getSyncTimestamp(), // BUGFIX (BUG PAYMENT REVERT): Wajib ada agar LWW selalu
      // memprioritaskan lokal atas versi cloud yang mungkin datang dengan timestamp berbeda
      // akibat konversi format Date oleh GAS/Sheets. Tanpa updatedAt, parseTime jatuh ke
      // `timestamp` yang bisa di-misparse → timeCloud > timeLocal → cloud menang → payment hilang.
    };
    // BUGFIX RACE CONDITION (BUG 1): Set isDbDirty SINKRON sebelum setDb agar
    // jika finalizeLogin fetch cloud selesai di saat yang sama, guard akan aktif
    // dan data payment baru tidak tertimpa. Jangan andalkan useEffect untuk ini
    // karena useEffect berjalan async (microtask queue) dan bisa terlambat.
    isDbDirty.current = true;
    // BUGFIX PAYMENT HILANG: Simpan payment baru ke localStorage SEGERA sebelum
    // sync debounce (2 detik) selesai, agar jika tab ditutup/browser crash,
    // data tidak hilang. setDb sudah memicu penyimpanan localStorage via useEffect,
    // tapi kita tambah guard eksplisit di sini untuk keamanan ekstra.
    setDb((p) => {
      const updated = { ...p, payments: [...p.payments, rec] };
      try { localStorage.setItem('ecg_db', JSON.stringify(updated)); } catch(e) {}
      return updated;
    });
    showToast(`Payment recorded`);
    setAmounts((p) => ({ ...p, [student.id]: '' }));
  };

  const localPrintPayment = () => {
    if (!selectedInvoice) return;
    const originalTitle = document.title;
    const safeName = selectedInvoice.studentName.replace(/[\s/\\?%*:|"<>-]/g, '_');
    document.title = `${safeName}_payment`;
    window.print();
    setTimeout(() => { document.title = originalTitle; }, 1000);
  };

  if (selectedInvoice) return (
    <div className="fixed inset-0 z-[100] bg-slate-50/95 backdrop-blur-md overflow-y-auto print:bg-white print:static print:block print:z-auto custom-scrollbar font-sans text-slate-900">
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[80%] h-[400px] bg-gradient-to-b from-blue-200/40 to-transparent blur-3xl pointer-events-none print:hidden" />
      
      <div className="w-full max-w-2xl mx-auto mt-6 mb-4 px-4 flex justify-between items-center relative z-10 print:hidden">
        <button className="flex items-center gap-2 text-slate-500 hover:text-slate-800 transition-colors font-medium text-sm bg-white/50 px-4 py-2 rounded-full border border-slate-200/50 shadow-sm" onClick={() => setSelectedInvoice(null)}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="flex gap-2">
          <button onClick={localPrintPayment} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Print PDF"><Printer size={16}/></button>
          <button onClick={() => {
             const safeName = selectedInvoice.studentName.replace(/[\s/\\?%*:|"<>-]/g, '_');
             downloadPNG('receipt-print', `${safeName}_payment`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Download PNG"><Download size={16}/></button>
          <button onClick={() => {
             const safeName = selectedInvoice.studentName.replace(/[\s/\\?%*:|"<>-]/g, '_');
             handleShareImage('receipt-print', `${safeName}_payment`, `Receipt for ${selectedInvoice.studentName}`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Share"><Share2 size={16}/></button>
        </div>
      </div>

      <div id="receipt-print" className="w-full max-w-2xl mx-auto bg-white rounded-none shadow-2xl overflow-hidden relative z-10 mb-12 print:m-0 print:shadow-none print:max-w-full print:rounded-none">
        <div className="bg-[#1A56DB] text-white p-6 sm:p-8 flex justify-between items-start">
          <div>
            <h1 className="text-2xl sm:text-3xl font-black tracking-tight">PAYMENT RECEIPT</h1>
            <p className="text-blue-200 font-mono mt-1 text-sm">{selectedInvoice.id}</p>
          </div>
          <div className="text-right">
            <p className="text-lg font-bold text-white leading-tight">English Club Gresik</p>
            <p className="text-blue-200 text-xs mt-1">Academic Suite</p>
          </div>
        </div>

        <div className="p-6 sm:p-8 pb-4">
          <div className="text-center mb-8 border-b border-slate-200 pb-8">
            <p className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">
              Tuition Fee {MONTHS[parseInt(selectedInvoice.month) - 1]} {selectedInvoice.year}
            </p>
            <p className="text-5xl font-black text-slate-900 tracking-tight">
              Rp {Number(selectedInvoice.amount).toLocaleString('id-ID')}
            </p>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">STUDENT INFO</p>
            <div className="bg-slate-50 border border-slate-200 p-5 shadow-sm">
              <p className="text-xl font-bold text-slate-800">{selectedInvoice.studentName}</p>
              <p className="text-sm font-medium text-slate-500 mt-1">{selectedInvoice.class} · {selectedInvoice.sessionGroup}</p>
            </div>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">TRANSACTION DETAILS</p>
            <div className="space-y-0 text-sm">
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Date & Time</span>
                <span className="font-semibold text-slate-800">
                  {/* FIX #5: safeDateDisplay agar konsisten dan bebas UTC-shift */}
                  {safeDateDisplay(selectedInvoice.date, 'en-US', { month: 'long', day: 'numeric', year: 'numeric' })} {selectedInvoice.time ? `• ${selectedInvoice.time}` : ''}
                </span>
              </div>
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Method</span>
                <span className="font-semibold text-slate-800">{selectedInvoice.method || 'Cash'}</span>
              </div>
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Receipt No.</span>
                <span className="font-mono font-semibold text-slate-800">{selectedInvoice.id}</span>
              </div>
              <div className="flex justify-between items-center py-3 px-2">
                <span className="font-medium text-slate-600">Status</span>
                <span className="text-lg font-black text-green-700">PAID</span>
              </div>
            </div>
          </div>

          <div className="text-center mt-12 mb-4">
            <p className="text-sm font-medium text-slate-600 italic">Thank you for your payment.</p>
            <p className="text-xs font-bold text-slate-500 mt-1">— Akhmad Akmal Rifqi</p>
          </div>
        </div>

        <div className="bg-[#1A56DB] p-6 sm:p-8 text-center">
          <p className="text-xs text-blue-100 font-medium mb-1">English Club Gresik • WA: 0897-327-11-12</p>
          <p className="text-[11px] text-blue-200 mb-3">Taman Anggrek Blok AB 05, Kedanyang, Kebomas, Gresik</p>
          <div className="border-t border-blue-400/30 pt-3">
            <p className="text-[11px] text-blue-200">
              This document serves as an official payment receipt
            </p>
          </div>
        </div>

      </div>
    </div>
  );

  return (
    <div className="space-y-6">
      
      {viewingStudentId && !selectedInvoice && (
        <CustomModal isOpen={true} onClose={() => setViewingStudentId(null)} title={`Receipts for ${activeStudents.find(s => s.id === viewingStudentId)?.name || 'Student'}`}>
           <div className="space-y-4">
              <table className="w-full text-left text-sm">
                 <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
                    <tr>
                       <th className="p-3 text-center w-12 text-gray-400">No.</th>
                       <th className="p-3 text-center">Receipt No</th>
                       <th className="p-3 text-center">Date</th>
                       <th className="p-3 text-center">Amount</th>
                       <th className="p-3 text-center">Actions</th>
                    </tr>
                 </thead>
                 <tbody className="divide-y divide-gray-800">
                    {db.payments.filter(p => p.studentId === viewingStudentId && Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid').map((p, index) => (
                       <tr key={p.id} className="hover:bg-[#0B0F19]">
                          <td className="p-3 text-center text-gray-500 font-medium">{index + 1}</td>
                          <td className="p-3 text-center font-mono text-gray-500 text-xs">{p.id}</td>
                          <td className="p-3 text-center text-white">{p.date}</td>
                          <td className="p-3 text-center text-green-400 font-bold">Rp {Number(p.amount).toLocaleString()}</td>
                          <td className="p-3 text-center flex justify-center gap-2">
                             <button onClick={() => setSelectedInvoice(p)} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="View Receipt"><FileText size={18}/></button>
                             <button onClick={() => softDelete('payments', p.id, 'Payment Receipt')} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Receipt"><Trash2 size={18}/></button>
                          </td>
                       </tr>
                    ))}
                 </tbody>
              </table>
           </div>
        </CustomModal>
      )}

      <div><h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Manajemen Pembayaran' : 'Payment Management'}</h2><p className="text-gray-400 text-sm">{language === 'id' ? 'Kelola pembayaran dan tagihan siswa.' : 'Track and manage student payments and billing.'}</p></div>
      
      <div className="bg-gradient-to-r from-orange-500/20 to-orange-600/10 border border-orange-500/30 p-5 rounded-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-5 shadow-lg relative overflow-hidden animation-fade-in">
          <div className="absolute -right-10 -top-10 w-40 h-40 bg-orange-500/10 blur-3xl rounded-full"></div>
          <div className="relative z-10 flex items-center gap-4">
              <div className="w-14 h-14 bg-white rounded-xl flex items-center justify-center font-black text-orange-600 text-xl shadow-md shrink-0 border-2 border-orange-500/20">
                  BNI
              </div>
              <div>
                  <p className="text-xs text-orange-400 font-bold uppercase tracking-widest mb-0.5">{language === 'id' ? 'Rekening Pembayaran Resmi' : 'Official Payment Account'}</p>
                  <p className="text-lg sm:text-xl font-bold text-white leading-tight">Akhmad Akmal Rifqi</p>
                  <p className="text-xs sm:text-sm text-gray-300 font-medium mt-0.5">Bank Negara Indonesia (BNI)</p>
              </div>
          </div>
          <div className="relative z-10 flex items-center justify-between w-full md:w-auto gap-4 bg-[#0B0F19]/80 backdrop-blur-sm px-5 py-3 rounded-xl border border-gray-700 shadow-inner">
              <span className="text-2xl sm:text-3xl font-mono font-black text-[#00D4FF] tracking-wider drop-shadow-md">0951837774</span>
              <button
                  onClick={() => {
                      navigator.clipboard.writeText('0951837774');
                      showToast(language === 'id' ? 'Nomor rekening BNI disalin!' : 'BNI Account number copied!');
                  }}
                  className="bg-[#00D4FF]/10 hover:bg-[#00D4FF]/20 text-[#00D4FF] p-2.5 rounded-lg transition-all active:scale-95 shadow-sm border border-[#00D4FF]/20"
                  title={language === 'id' ? 'Salin Rekening' : 'Copy Account Number'}
              >
                  <Copy size={20} />
              </button>
          </div>
      </div>

      {/* ─── REVENUE SUMMARY CARDS ─── */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">

        {/* Card 1: Total Revenue (Collected) */}
        <div className="relative overflow-hidden rounded-2xl border border-[#00D4FF]/30 bg-gradient-to-br from-[#00D4FF]/15 via-[#0B1A2E] to-[#151B26] shadow-[0_0_30px_rgba(0,212,255,0.12)] p-5">
          {/* Glow orb */}
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full bg-[#00D4FF]/10 blur-2xl pointer-events-none" />
          <div className="relative z-10 flex flex-col h-full gap-3">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold text-[#00D4FF] uppercase tracking-widest">{language === 'id' ? 'Total Pendapatan' : 'Total Revenue'}</span>
              <div className="flex items-center justify-center w-9 h-9 rounded-xl bg-[#00D4FF]/20 border border-[#00D4FF]/30">
                <DollarSign size={18} className="text-[#00D4FF]" />
              </div>
            </div>
            <div>
              <p className="text-[11px] text-gray-500 mb-0.5">{MONTHS[month - 1]} {year}</p>
              <p className="text-3xl font-black text-white tracking-tight leading-none">Rp {totalRevenue.toLocaleString()}</p>
            </div>
            {/* Collection rate progress bar */}
            <div className="mt-auto pt-2">
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-[11px] text-gray-400">{language === 'id' ? 'Tingkat Koleksi' : 'Collection Rate'}</span>
                <span className="text-[11px] font-bold text-[#00D4FF]">{revenueStats.collectionRate}%</span>
              </div>
              <div className="w-full h-1.5 bg-[#00D4FF]/10 rounded-full overflow-hidden">
                <div
                  className="h-full rounded-full bg-gradient-to-r from-[#00D4FF] to-[#2563EB] transition-all duration-700"
                  style={{ width: `${revenueStats.collectionRate}%` }}
                />
              </div>
              <div className="flex gap-3 mt-2.5 text-[11px] text-gray-400">
                <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-400 inline-block"/>{revenueStats.paidCount} {language === 'id' ? 'Lunas' : 'Paid'}</span>
                <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-amber-400 inline-block"/>{revenueStats.partialCount} {language === 'id' ? 'Sebagian' : 'Partial'}</span>
                <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-rose-400 inline-block"/>{revenueStats.unpaidCount} {language === 'id' ? 'Belum Bayar' : 'Unpaid'}</span>
              </div>
            </div>
          </div>
        </div>

        {/* Card 2: Outstanding / Belum Terbayar */}
        <div className="relative overflow-hidden rounded-2xl border border-rose-500/30 bg-gradient-to-br from-rose-500/10 via-[#0B1A2E] to-[#151B26] shadow-[0_0_30px_rgba(239,68,68,0.10)] p-5">
          {/* Glow orb */}
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full bg-rose-500/10 blur-2xl pointer-events-none" />
          <div className="relative z-10 flex flex-col h-full gap-3">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-bold text-rose-400 uppercase tracking-widest">{language === 'id' ? 'Belum Terbayar' : 'Outstanding'}</span>
              <div className="flex items-center justify-center w-9 h-9 rounded-xl bg-rose-500/15 border border-rose-500/30">
                <AlertCircle size={18} className="text-rose-400" />
              </div>
            </div>
            <div>
              <p className="text-[11px] text-gray-500 mb-0.5">{MONTHS[month - 1]} {year}</p>
              <p className={`text-3xl font-black tracking-tight leading-none ${revenueStats.totalOutstanding > 0 ? 'text-rose-400' : 'text-emerald-400'}`}>
                {revenueStats.totalOutstanding > 0 ? `Rp ${revenueStats.totalOutstanding.toLocaleString()}` : (language === 'id' ? 'Semua Lunas! 🎉' : 'All Clear! 🎉')}
              </p>
            </div>
            {/* Breakdown */}
            <div className="mt-auto pt-2 space-y-2">
              <div className="flex items-center justify-between bg-[#0B0F19]/60 rounded-lg px-3 py-2 border border-gray-800">
                <span className="text-[11px] text-gray-400">{language === 'id' ? 'Total Tagihan' : 'Total Billed'}</span>
                <span className="text-[11px] font-bold text-gray-200">Rp {revenueStats.totalBilled.toLocaleString()}</span>
              </div>
              <div className="flex items-center justify-between bg-rose-500/5 rounded-lg px-3 py-2 border border-rose-500/20">
                <div className="flex items-center gap-1.5">
                  <span className="w-1.5 h-4 rounded-full bg-rose-500 inline-block"/>
                  <span className="text-[11px] text-rose-300">{language === 'id' ? 'Kekurangan' : 'Deficit'}</span>
                </div>
                <span className="text-[11px] font-black text-rose-400">
                  {revenueStats.totalOutstanding > 0 ? `− Rp ${revenueStats.totalOutstanding.toLocaleString()}` : '—'}
                </span>
              </div>
            </div>
          </div>
        </div>

      </div>
      {/* ─── END REVENUE SUMMARY CARDS ─── */}

      {/* ─── PAYMENT QUICK OVERVIEW PANEL ─── */}
      {(() => {
        // Compute per-student status for ALL active students (not just paginated)
        const allStudentStatuses = activeStudents.map((s) => {
          const pays = db.payments.filter(
            // BUGFIX: trim+toLowerCase agar tahan terhadap transformasi casing/whitespace dari GAS
            (p) => String(p.studentId).trim().toLowerCase() === String(s.id).trim().toLowerCase() && Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid'
          );
          const totalPaid = pays.reduce((sum, p) => sum + Number(p.amount), 0);
          const target = getStudentTarget(s);
          const lastPay = pays.sort((a, b) => (a.date > b.date ? -1 : 1))[0];
          let status: 'Paid' | 'Unpaid' | 'Partial' | 'No Target' | 'Debt' | 'Deposit' = 'Unpaid';
          if (s.paymentPlan === 'Free') status = 'No Target';
          else if (target === 0 && totalPaid === 0) status = 'No Target';
          else if (s.paymentPlan === 'Per Visit') {
            const bal = totalPaid - target;
            if (bal < 0) status = 'Debt';
            else if (bal > 0) status = 'Deposit';
            else if (bal === 0 && target > 0) status = 'Paid';
            else status = 'No Target';
          } else {
            if (totalPaid >= target && target > 0) status = 'Paid';
            else if (totalPaid > 0) status = 'Partial';
          }
          return { s, pays, totalPaid, target, lastPay, status };
        });

        const paidCount = allStudentStatuses.filter(x => x.status === 'Paid' || x.status === 'Deposit').length;
        const unpaidCount = allStudentStatuses.filter(x => x.status === 'Unpaid' || x.status === 'Debt').length;
        const partialCount = allStudentStatuses.filter(x => x.status === 'Partial').length;
        const totalActive = allStudentStatuses.filter(x => x.status !== 'No Target').length;
        const paidPct = totalActive > 0 ? Math.round((paidCount / totalActive) * 100) : 0;

        const filtered = overviewFilter === 'all' ? allStudentStatuses
          : overviewFilter === 'paid' ? allStudentStatuses.filter(x => x.status === 'Paid' || x.status === 'Deposit')
          : overviewFilter === 'unpaid' ? allStudentStatuses.filter(x => x.status === 'Unpaid' || x.status === 'Debt')
          : allStudentStatuses.filter(x => x.status === 'Partial');

        const statusColor = {
          Paid: 'text-emerald-400', Deposit: 'text-cyan-400',
          Unpaid: 'text-rose-400', Debt: 'text-red-400',
          Partial: 'text-amber-400', 'No Target': 'text-gray-500',
        };
        const statusBg = {
          Paid: 'bg-emerald-500/10 border-emerald-500/30', Deposit: 'bg-cyan-500/10 border-cyan-500/30',
          Unpaid: 'bg-rose-500/10 border-rose-500/30', Debt: 'bg-red-500/10 border-red-500/30',
          Partial: 'bg-amber-500/10 border-amber-500/30', 'No Target': 'bg-gray-800/50 border-gray-700',
        };

        return (
          <Card className="p-0 overflow-hidden border border-gray-800">
            {/* Header row — always visible */}
            <div
              className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 px-5 py-4 bg-[#0A0E17] cursor-pointer select-none"
              onClick={() => setShowQuickOverview(v => !v)}
            >
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-violet-500/20 border border-violet-500/30 flex items-center justify-center shrink-0">
                  <DollarSign size={18} className="text-violet-400" />
                </div>
                <div>
                  <p className="text-sm font-bold text-white">{language === 'id' ? 'Ringkasan Status Bayar — Cepat' : 'Quick Payment Overview'}</p>
                  <p className="text-xs text-gray-400">{MONTHS[month - 1]} {year} · {language === 'id' ? 'Klik untuk' : 'Click to'} {showQuickOverview ? (language === 'id' ? 'tutup' : 'collapse') : (language === 'id' ? 'buka' : 'expand')}</p>
                </div>
              </div>
              {/* Mini stat chips */}
              <div className="flex items-center gap-2 flex-wrap">
                <span className="px-3 py-1 rounded-full bg-emerald-500/15 border border-emerald-500/30 text-emerald-400 text-xs font-bold">✓ {paidCount} {language === 'id' ? 'Lunas' : 'Paid'}</span>
                <span className="px-3 py-1 rounded-full bg-amber-500/15 border border-amber-500/30 text-amber-400 text-xs font-bold">◑ {partialCount} {language === 'id' ? 'Parsial' : 'Partial'}</span>
                <span className="px-3 py-1 rounded-full bg-rose-500/15 border border-rose-500/30 text-rose-400 text-xs font-bold">✗ {unpaidCount} {language === 'id' ? 'Belum' : 'Unpaid'}</span>
                <span className="px-3 py-1 rounded-full bg-[#151B26] border border-gray-700 text-gray-300 text-xs font-bold">{paidPct}% {language === 'id' ? 'terkumpul' : 'collected'}</span>
                <span className="text-gray-600 text-lg ml-1">{showQuickOverview ? '▲' : '▼'}</span>
              </div>
            </div>

            {showQuickOverview && (
              <div className="border-t border-gray-800">
                {/* Progress bar */}
                <div className="px-5 pt-4 pb-2">
                  <div className="flex justify-between text-[11px] text-gray-400 mb-1">
                    <span>{language === 'id' ? 'Tingkat pelunasan' : 'Collection rate'}</span>
                    <span className="font-bold text-white">{paidPct}%</span>
                  </div>
                  <div className="w-full h-2 bg-gray-800 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-gradient-to-r from-emerald-500 to-cyan-400 rounded-full transition-all duration-500"
                      style={{ width: `${paidPct}%` }}
                    />
                  </div>
                  <div className="flex justify-between text-[10px] text-gray-600 mt-1">
                    <span>{paidCount} {language === 'id' ? 'dari' : 'of'} {totalActive} {language === 'id' ? 'siswa aktif' : 'active students'}</span>
                    <span>{unpaidCount + partialCount} {language === 'id' ? 'belum lunas' : 'outstanding'}</span>
                  </div>
                </div>

                {/* Filter tabs */}
                <div className="px-5 py-3 flex gap-2 flex-wrap border-t border-gray-800/60">
                  {(['all', 'paid', 'unpaid', 'partial'] as const).map(f => (
                    <button
                      key={f}
                      onClick={() => setOverviewFilter(f)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all border ${overviewFilter === f
                        ? f === 'paid' ? 'bg-emerald-500/20 border-emerald-500/50 text-emerald-300'
                        : f === 'unpaid' ? 'bg-rose-500/20 border-rose-500/50 text-rose-300'
                        : f === 'partial' ? 'bg-amber-500/20 border-amber-500/50 text-amber-300'
                        : 'bg-violet-500/20 border-violet-500/50 text-violet-300'
                        : 'bg-[#0A0E17] border-gray-700 text-gray-400 hover:border-gray-500'
                      }`}
                    >
                      {f === 'all' ? `${language === 'id' ? 'Semua' : 'All'} (${allStudentStatuses.length})`
                        : f === 'paid' ? `✓ ${language === 'id' ? 'Lunas' : 'Paid'} (${paidCount})`
                        : f === 'unpaid' ? `✗ ${language === 'id' ? 'Belum Bayar' : 'Unpaid'} (${unpaidCount})`
                        : `◑ ${language === 'id' ? 'Parsial' : 'Partial'} (${partialCount})`}
                    </button>
                  ))}
                </div>

                {/* Cards grid */}
                <div className="px-4 pb-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3 max-h-[480px] overflow-y-auto">
                  {filtered.length === 0 && (
                    <div className="col-span-full py-10 text-center text-gray-500 text-sm">
                      {language === 'id' ? 'Tidak ada siswa dengan status ini.' : 'No students with this status.'}
                    </div>
                  )}
                  {filtered.map(({ s, pays, totalPaid, target, lastPay, status }) => {
                    const balance = totalPaid - target;
                    const pct = target > 0 ? Math.min(100, Math.round((totalPaid / target) * 100)) : 0;
                    return (
                      <div
                        key={s.id}
                        className={`rounded-xl border p-3.5 flex flex-col gap-2 transition-all hover:-translate-y-0.5 hover:shadow-lg ${statusBg[status] || 'bg-gray-800/50 border-gray-700'}`}
                      >
                        {/* Student name + plan badge */}
                        <div className="flex items-start justify-between gap-2">
                          <div className="min-w-0">
                            <p className="text-sm font-bold text-white truncate">{s.name}</p>
                            <p className="text-[11px] text-gray-400">{s.class}</p>
                          </div>
                          <span className={`shrink-0 text-[10px] px-2 py-0.5 rounded-full border font-bold uppercase tracking-wide ${s.paymentPlan === 'Per Visit' ? 'bg-purple-500/20 text-purple-300 border-purple-500/40' : s.paymentPlan === 'Free' ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' : 'bg-blue-500/20 text-blue-300 border-blue-500/40'}`}>
                            {s.paymentPlan === 'Per Visit' ? 'Per Visit' : s.paymentPlan === 'Free' ? '🎁 Free' : 'Monthly'}
                          </span>
                        </div>

                        {/* Status badge */}
                        <div className="flex items-center gap-2">
                          <Badge status={status} />
                          {pays.length > 0 && (
                            <span className="text-[10px] text-gray-500">{pays.length}x {language === 'id' ? 'transaksi' : 'transaction(s)'}</span>
                          )}
                        </div>

                        {/* Amount paid / target */}
                        {target > 0 && (
                          <>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-400">{language === 'id' ? 'Dibayar' : 'Paid'}</span>
                              <span className={`font-bold ${statusColor[status] || 'text-gray-300'}`}>Rp {totalPaid.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-400">{language === 'id' ? 'Target' : 'Target'}</span>
                              <span className="text-gray-300 font-medium">Rp {target.toLocaleString()}</span>
                            </div>
                            {/* Mini progress bar */}
                            <div className="w-full h-1.5 bg-gray-800 rounded-full overflow-hidden">
                              <div
                                className={`h-full rounded-full transition-all duration-300 ${status === 'Paid' || status === 'Deposit' ? 'bg-emerald-400' : status === 'Partial' ? 'bg-amber-400' : 'bg-rose-500'}`}
                                style={{ width: `${pct}%` }}
                              />
                            </div>
                            {s.paymentPlan === 'Per Visit' && (
                              <div className="flex justify-between text-xs">
                                <span className="text-gray-400">{balance < 0 ? (language === 'id' ? 'Sisa Tagihan' : 'Remaining') : (language === 'id' ? 'Saldo/Deposit' : 'Deposit')}</span>
                                <span className={`font-bold ${balance < 0 ? 'text-red-400' : 'text-cyan-400'}`}>Rp {Math.abs(balance).toLocaleString()}</span>
                              </div>
                            )}
                          </>
                        )}

                        {/* Payment history — tanggal dan nominal */}
                        {pays.length > 0 && (
                          <div className="mt-1 border-t border-gray-700/50 pt-2 flex flex-col gap-1">
                            <p className="text-[10px] text-gray-500 uppercase tracking-wider font-bold mb-0.5">{language === 'id' ? 'Riwayat Bayar' : 'Payment History'}</p>
                            {pays.sort((a, b) => (a.date > b.date ? 1 : -1)).map((p, i) => (
                              <div key={i} className="flex justify-between items-center text-[11px]">
                                <span className="text-gray-400">{p.date ? p.date.split('T')[0] : '-'} <span className="text-gray-600">·</span> <span className="text-gray-500">{p.method || 'Cash'}</span></span>
                                <span className="text-white font-semibold">Rp {Number(p.amount).toLocaleString()}</span>
                              </div>
                            ))}
                          </div>
                        )}

                        {/* Belum bayar sama sekali */}
                        {pays.length === 0 && target > 0 && (
                          <div className="mt-1 border-t border-gray-700/50 pt-2">
                            <p className="text-[11px] text-rose-400/80 italic">{language === 'id' ? 'Belum ada pembayaran.' : 'No payment recorded yet.'}</p>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </Card>
        );
      })()}
      {/* ─── END PAYMENT QUICK OVERVIEW ─── */}

      <Card className="p-0 overflow-hidden flex flex-col">
        {/* Unified Filter Row: Month, Year, Session, Search, Level, Class - all in one Card */}
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
          <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={month} onChange={(e) => setMonth(Number(e.target.value))}>
            {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
          </select>
          <input type="number" className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={year} onChange={(e) => setYear(Number(e.target.value))} />
          <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={sessionGroup} onChange={(e) => setSessionGroup(e.target.value)}>
            <option value="All Sessions">All Sessions</option>
            {SESSIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input type="text" placeholder="Search students..." className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
          <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterLevel} onChange={(e) => { setFilterLevel(e.target.value); setFilterClass(''); }}>
            <option value="">All Levels</option>
            {LEVELS.map((l) => <option key={l} value={l}>{l}</option>)}
          </select>
          <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterClass} onChange={(e) => setFilterClass(e.target.value)} disabled={!filterLevel}>
            <option value="">All Classes</option>
            {filterLevel && CLASS_MAPPING[filterLevel]?.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        
        {/* ── MOBILE CARD LIST (< sm) ── */}
        <div className="sm:hidden divide-y divide-gray-800">
          {paginatedData.length === 0 ? (
            <EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." />
          ) : paginatedData.map((s, index) => {
            const studentPayments = db.payments.filter((p) => String(p.studentId).trim().toLowerCase() === String(s.id).trim().toLowerCase() && Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid');
            const totalPaid = studentPayments.reduce((sum, p) => sum + Number(p.amount), 0);
            const target = getStudentTarget(s);
            let status: any = 'Unpaid';
            if (s.paymentPlan === 'Free') status = 'No Target';
            else if (target === 0 && totalPaid === 0) status = 'No Target';
            else if (totalPaid >= target && target > 0) status = 'Paid';
            else if (totalPaid > 0) status = 'Partial';
            const balance = totalPaid - target;
            if (s.paymentPlan === 'Per Visit') {
              if (balance < 0) status = 'Debt';
              else if (balance > 0) status = 'Deposit';
              else if (balance === 0 && target > 0) status = 'Paid';
              else status = 'No Target';
            }
            const count = attendanceThisMonth[s.id] || 0;
            return (
              <div key={s.id} className="p-4 space-y-3 hover:bg-[#0B0F19]/60 transition-colors">
                {/* Row 1: name + status badge */}
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="flex flex-wrap items-center gap-1">
                      <span className="font-bold text-white text-sm">{s.name}</span>
                      <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} />
                    </div>
                    <p className="text-[11px] text-gray-500 mt-0.5">{s.class} · <span className={`font-semibold ${s.paymentPlan === 'Per Visit' ? 'text-purple-300' : s.paymentPlan === 'Free' ? 'text-emerald-300' : 'text-blue-300'}`}>{s.paymentPlan}</span> · {count}x hadir</p>
                  </div>
                  <Badge status={status} />
                </div>
                {/* Row 2: balance info */}
                <div className="bg-[#0B0F19] rounded-lg px-3 py-2 text-xs flex justify-between items-center">
                  {s.paymentPlan === 'Free' ? (
                    <span className="text-emerald-400 font-bold">🎁 Complimentary</span>
                  ) : s.paymentPlan === 'Per Visit' ? (
                    <>
                      <span className="text-gray-400">Balance</span>
                      <span className={`font-bold ${balance < 0 ? 'text-red-400' : balance > 0 ? 'text-cyan-400' : 'text-gray-500'}`}>Rp {balance.toLocaleString()}</span>
                    </>
                  ) : (
                    <>
                      <span className="text-gray-400">Paid / Target</span>
                      <span className={`font-bold ${totalPaid >= target && target > 0 ? 'text-green-400' : totalPaid > 0 ? 'text-yellow-400' : 'text-gray-500'}`}>Rp {totalPaid.toLocaleString()} <span className="text-gray-600 font-normal">/ Rp {target.toLocaleString()}</span></span>
                    </>
                  )}
                </div>
                {/* Row 3: payment input + action buttons */}
                <div className="flex flex-col gap-2">
                  {s.paymentPlan !== 'Free' && !(s.paymentPlan === 'Monthly' && status === 'Paid') && (
                    <div className="flex items-center bg-[#0B0F19] rounded-lg border border-gray-700 overflow-hidden">
                      <select className="bg-transparent border-none px-2 py-2 text-white text-xs focus:ring-0 focus:outline-none w-[72px] cursor-pointer" value={methods[s.id] || 'Cash'} onChange={(e) => setMethods((p) => ({ ...p, [s.id]: e.target.value }))}>
                        <option value="Cash">Cash</option>
                        <option value="Transfer">Transfer</option>
                      </select>
                      <div className="w-px h-5 bg-gray-700"/>
                      <input type="number" className="bg-transparent border-none px-2 py-2 text-white text-xs focus:ring-0 focus:outline-none flex-1 min-w-0 placeholder-gray-600" placeholder="Custom Rp" value={amounts[s.id] || ''} onChange={(e) => setAmounts((p) => ({ ...p, [s.id]: e.target.value }))} />
                      <div className="w-px h-5 bg-gray-700"/>
                      <button onClick={() => handleRecordInline(s, undefined)} disabled={!amounts[s.id]} className="bg-[#00D4FF]/10 text-[#00D4FF] hover:bg-[#00D4FF]/20 px-3 py-2 text-xs font-bold transition-colors disabled:opacity-50">Add</button>
                      <div className="w-px h-5 bg-gray-700"/>
                      <button onClick={() => handleRecordInline(s, 25000)} className="bg-green-500/10 text-green-400 hover:bg-green-500 hover:text-white px-3 py-2 text-xs font-bold transition-all whitespace-nowrap">+25k</button>
                    </div>
                  )}
                  <div className="flex gap-2">
                    {studentPayments.length > 0 && (
                      <button onClick={() => setViewingStudentId(s.id)} className="flex-1 flex items-center justify-center gap-1.5 bg-[#151B26] border border-gray-700 hover:border-blue-500/50 px-3 py-2 rounded-lg transition-colors text-xs text-gray-300">
                        <FileText size={13} className="text-blue-400" /><span>Receipts ({studentPayments.length})</span>
                      </button>
                    )}
                    {s.paymentPlan === 'Per Visit' && (
                      <button onClick={() => { const monthPrefix = `${year}-${String(month).padStart(2, '0')}`; const presentCount = db.studentAttendance.filter(a => a.studentId === s.id && a.date.startsWith(monthPrefix) && a.status === 'Present').length; const balanceText = balance < 0 ? `Remaining: Rp ${Math.abs(balance).toLocaleString()}` : balance > 0 ? `Deposit: Rp ${balance.toLocaleString()}` : 'Paid'; const message = `Hi, payment summary for ${s.name} — ${MONTHS[month - 1]} ${year}:\n\nAttended: ${presentCount}x\nPaid: Rp ${totalPaid.toLocaleString()}\nTarget: Rp ${target.toLocaleString()}\n${balanceText}\n\nThank you.`; window.open(`https://wa.me/${normalizeWhatsapp(s.whatsapp)}?text=${encodeURIComponent(message)}`, '_blank'); }} className="flex-1 flex items-center justify-center gap-1.5 bg-[#151B26] border border-gray-700 hover:border-green-500/50 hover:bg-green-500/10 px-3 py-2 rounded-lg transition-colors text-xs text-gray-300">
                        <MessageCircle size={13} className="text-green-400" /><span>Bill WA</span>
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* ── DESKTOP TABLE (≥ sm) ── */}
        <div className="hidden sm:block overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">Student</th><th className="p-4 text-center">Class & Plan</th><th className="p-4 text-center">Attended<br/><span className="text-[10px] font-normal text-gray-500 normal-case tracking-normal">This Month</span></th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Balance / Target</th><th className="p-4 text-center">Actions</th></tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {paginatedData.map((s, index) => {
              // BUGFIX (BUG PAYMENT REVERT): Gunakan trim().toLowerCase() agar filter tetap
              // cocok meski GAS/Sheets mengubah casing atau menambah whitespace pada studentId.
              const studentPayments = db.payments.filter((p) => String(p.studentId).trim().toLowerCase() === String(s.id).trim().toLowerCase() && Number(p.month) === Number(month) && String(p.year) === String(year) && p.status === 'Paid');
              const totalPaid = studentPayments.reduce((sum, p) => sum + Number(p.amount), 0);
              const target = getStudentTarget(s);
              
              let status = 'Unpaid';
              if (s.paymentPlan === 'Free') status = 'No Target';
              else if (target === 0 && totalPaid === 0) status = 'No Target';
              else if (totalPaid >= target && target > 0) status = 'Paid';
              else if (totalPaid > 0) status = 'Partial';

              const balance = totalPaid - target;
              if (s.paymentPlan === 'Per Visit') {
                if (balance < 0) status = 'Debt';
                else if (balance > 0) status = 'Deposit';
                else if (balance === 0 && target > 0) status = 'Paid';
                else status = 'No Target';
              }

              return (
                <tr key={s.id} className="hover:bg-[#0B0F19]">
                  <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                  <td className="p-4 text-center text-white font-medium"><div className="flex items-center justify-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /></div></td>
                  <td className="p-4 text-center leading-tight">
                    <span className="font-bold text-gray-300">{s.class}</span> <br/>
                    <span className={`text-[11px] px-2 py-0.5 rounded font-semibold uppercase tracking-wide mt-1.5 inline-block border ${s.paymentPlan === 'Per Visit' ? 'bg-purple-500/20 text-purple-300 border-purple-500/50 shadow-[0_0_10px_rgba(168,85,247,0.3)]' : s.paymentPlan === 'Free' ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/50 shadow-[0_0_10px_rgba(16,185,129,0.3)]' : 'bg-blue-500/20 text-blue-300 border-blue-500/50 shadow-[0_0_10px_rgba(59,130,246,0.3)]'}`}>
                      {s.paymentPlan === 'Per Visit' ? 'Per Visit' : s.paymentPlan === 'Free' ? '🎁 Free' : 'Monthly'}
                    </span>
                  </td>
                  <td className="p-4 text-center">
                    {(() => {
                      const count = attendanceThisMonth[s.id] || 0;
                      const isPerVisit = s.paymentPlan === 'Per Visit';
                      return (
                        <span className={`inline-flex items-center justify-center min-w-[2rem] px-2 py-1 rounded-full text-sm font-bold ${
                          count === 0 ? 'bg-gray-700/50 text-gray-500' :
                          isPerVisit ? 'bg-purple-500/20 text-purple-300 border border-purple-500/40' :
                          'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                        }`}>
                          {count}x
                        </span>
                      );
                    })()}
                  </td>
                  <td className="p-4 text-center"><Badge status={status} /></td>
                  <td className="p-4 text-center">
                     {s.paymentPlan === 'Free' ? (
                       <span className="text-emerald-400 text-xs font-bold">Complimentary</span>
                     ) : s.paymentPlan === 'Per Visit' ? (
                       <>
                         <span className={`font-bold ${balance < 0 ? 'text-red-400' : balance > 0 ? 'text-cyan-400' : 'text-gray-500'}`}>Rp {balance.toLocaleString()}</span>
                         <br/>
                         <span className="text-gray-500 text-[11px]">Target: Rp {target.toLocaleString()}</span>
                       </>
                     ) : (
                       <>
                         <span className={`font-bold ${totalPaid >= target && target > 0 ? 'text-green-400' : totalPaid > 0 ? 'text-yellow-400' : 'text-gray-500'}`}>Rp {totalPaid.toLocaleString()}</span>
                         <span className="text-gray-500 text-xs mx-1">/</span>
                         <span className="text-gray-400 text-xs font-semibold">Rp {target.toLocaleString()}</span>
                       </>
                     )}
                  </td>
                  <td className="p-4 align-middle">
                    <div className="flex items-center justify-center gap-3 flex-wrap xl:flex-nowrap">
                      {s.paymentPlan === 'Free' ? (
                        <span className="text-xs text-emerald-400 font-semibold bg-emerald-500/10 border border-emerald-500/20 px-3 py-1.5 rounded-lg">🎁 No billing</span>
                      ) : (s.paymentPlan === 'Per Visit' || !(s.paymentPlan === 'Monthly' && status === 'Paid')) && (
                         <div className="flex items-center bg-[#0B0F19] rounded-lg border border-gray-700 overflow-hidden shadow-inner shrink-0 transition-all focus-within:border-[#00D4FF]/60 focus-within:ring-1 focus-within:ring-[#00D4FF]/20">
                           <select
                              className="bg-transparent border-none px-2.5 py-2 text-white text-[11px] focus:ring-0 focus:outline-none w-[80px] cursor-pointer appearance-none"
                              value={methods[s.id] || 'Cash'}
                              onChange={(e) => setMethods((p) => ({ ...p, [s.id]: e.target.value }))}
                              style={{ backgroundImage: `url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%236b7280' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e")`, backgroundPosition: `right 0.2rem center`, backgroundRepeat: `no-repeat`, backgroundSize: `1.2em 1.2em`, paddingRight: '1.2rem' }}
                           >
                              <option value="Cash" className="bg-[#0B0F19]">Cash</option>
                              <option value="Transfer" className="bg-[#0B0F19]">Transfer</option>
                           </select>
                           <div className="w-px h-5 bg-gray-700"></div>
                           <input
                              type="number"
                              className="bg-transparent border-none px-3 py-2 text-white text-[11px] focus:ring-0 focus:outline-none w-[90px] placeholder-gray-600"
                              placeholder="Custom Rp"
                              value={amounts[s.id] || ''}
                              onChange={(e) => setAmounts((p) => ({ ...p, [s.id]: e.target.value }))}
                           />
                           <div className="w-px h-5 bg-gray-700"></div>
                           <button
                              onClick={() => handleRecordInline(s, undefined)}
                              disabled={!amounts[s.id]}
                              className="bg-[#00D4FF]/10 text-[#00D4FF] hover:bg-[#00D4FF]/20 px-4 py-2 text-[11px] font-bold transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                           >
                              Add
                           </button>
                           <div className="w-px h-5 bg-gray-700"></div>
                           <button
                              onClick={() => handleRecordInline(s, 25000)}
                              className="bg-green-500/10 text-green-400 hover:bg-green-500 hover:text-white px-4 py-2 text-[11px] font-bold transition-all whitespace-nowrap"
                           >
                              + 25k
                           </button>
                         </div>
                      )}
                      
                      {studentPayments.length > 0 && (
                         <button
                           onClick={() => setViewingStudentId(s.id)}
                           className="flex items-center gap-1.5 bg-[#151B26] border border-gray-700 hover:border-blue-500/50 hover:bg-[#1A2234] px-3 py-2 rounded-lg transition-colors text-[11px] text-gray-300 shadow-sm shrink-0"
                         >
                           <FileText size={14} className="text-blue-400" />
                           <span>Receipts</span>
                           <span className="bg-blue-500/20 text-blue-400 text-[11px] px-1.5 py-0.5 rounded-full font-bold ml-1">{studentPayments.length}</span>
                         </button>
                      )}

                      {s.paymentPlan === 'Per Visit' && (
                         <button
                           onClick={() => {
                              const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
                              const presentCount = db.studentAttendance.filter(a => a.studentId === s.id && a.date.startsWith(monthPrefix) && a.status === 'Present').length;
                              const balanceText = balance < 0
                                 ? `Remaining Balance: Rp ${Math.abs(balance).toLocaleString()}`
                                 : balance > 0
                                 ? `Deposit/Credit: Rp ${balance.toLocaleString()}`
                                 : 'Status: Paid';
                              const message = `Hi, here is the payment summary for ${s.name} — ${MONTHS[month - 1]} ${year}:\n\nAttended: ${presentCount}x\nAmount Paid: Rp ${totalPaid.toLocaleString()}\nTarget: Rp ${target.toLocaleString()}\n${balanceText}\n\nThank you.`;
                              const phone = normalizeWhatsapp(s.whatsapp);
                              window.open(`https://wa.me/${phone}?text=${encodeURIComponent(message)}`, '_blank');
                           }}
                           className="flex items-center gap-1.5 bg-[#151B26] border border-gray-700 hover:border-green-500/50 hover:bg-green-500/10 px-3 py-2 rounded-lg transition-colors text-[11px] text-gray-300 shadow-sm shrink-0"
                         >
                           <MessageCircle size={14} className="text-green-400" />
                           <span>Bill via WA</span>
                         </button>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
            {paginatedData.length === 0 && (
               <tr><td colSpan={7}><EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." /></td></tr>
            )}
          </tbody>
        </table>
        </div>

        {/* NEW: Pagination Footer */}
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
           <div className="flex items-center gap-2">
             <span>Show</span>
             <select 
               value={rowsPerPage} 
               onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} 
               className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer"
             >
               <option value={10}>10</option>
               <option value={20}>20</option>
               <option value={30}>30</option>
               <option value="All">All</option>
             </select>
             <span>entries {filteredStudents.length > 0 && `(Total: ${filteredStudents.length})`}</span>
           </div>
           
           {!isAll && totalPages > 1 && (
             <div className="flex items-center gap-2">
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
               <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
             </div>
           )}
        </div>
      </Card>
    </div>
  );
}

function HistoryReportsModule({ db, setDb, showToast, handlePrint, user, handleShareImage, downloadPNG }) {
  const [view, setView] = useState('directory');
  const [dirType, setDirType] = useState('student');
  const [selectedId, setSelectedId] = useState(null);

  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearchTerm = useDebounce(searchTerm, 300);
  const [filterSession, setFilterSession] = useState('');
  
  // NEW: State for Filters & Pagination
  const [filterLevel, setFilterLevel] = useState('');
  const [filterClass, setFilterClass] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);
  // Fix 2: Month/Year filter for directory
  const [dirFilterMonth, setDirFilterMonth] = useState<number | string>('All');
  const [dirFilterYear, setDirFilterYear] = useState(new Date().getFullYear());
  
  const [currentTeacherComment, setCurrentTeacherComment] = useState('');

  // Report period: driven by month+year dropdowns; startDate/endDate auto-computed
  const [reportPickerMonth, setReportPickerMonth] = useState<number>(() => new Date().getMonth() + 1);
  const [reportPickerYear, setReportPickerYear] = useState<number>(() => new Date().getFullYear());

  const startDate = useMemo(() => {
    return `${reportPickerYear}-${String(reportPickerMonth).padStart(2, '0')}-01`;
  }, [reportPickerMonth, reportPickerYear]);

  const endDate = useMemo(() => {
    // Last day of the selected month
    const lastDay = new Date(reportPickerYear, reportPickerMonth, 0).getDate();
    return `${reportPickerYear}-${String(reportPickerMonth).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  }, [reportPickerMonth, reportPickerYear]);

  useEffect(() => {
    setCurrentPage(1);
  }, [debouncedSearchTerm, filterSession, filterLevel, filterClass, dirType, rowsPerPage, dirFilterMonth, dirFilterYear]);

  const openProfile = (id, type) => {
    setSelectedId(id);
    setView(type === 'student' ? 'studentProfile' : 'tutorProfile');
    // Reset to current month/year whenever a new profile is opened
    const now = new Date();
    setReportPickerMonth(now.getMonth() + 1);
    setReportPickerYear(now.getFullYear());
  };

  const student = view === 'studentProfile' ? db.students.find((s) => s.id === selectedId) : null;

  useEffect(() => {
    if (student) setCurrentTeacherComment(student.teacherComment || '');
  }, [student]);

  // --- MOVED HOOKS TO TOP LEVEL ---
  const filteredStudents = useMemo(() => {
     return sortStudentsLogically(db.students.filter((s) => {
       const matchSearch = (s.name || '').toLowerCase().includes((debouncedSearchTerm || '').toLowerCase());
       const matchSession = filterSession ? getStudentSession(s) === filterSession : true;
       const matchTutor = user?.role === 'tutor' ? parseSessions(user.teachingSession).includes(getStudentSession(s)) : true;
       const matchLevel = filterLevel ? s.level === filterLevel : true;
       const matchClass = filterClass ? s.class === filterClass : true;
       // Fix 2: Month/Year filter — check if student has any attendance record in selected period
       let matchPeriod = true;
       if (dirFilterMonth !== 'All') {
         const prefix = `${dirFilterYear}-${String(dirFilterMonth).padStart(2, '0')}`;
         matchPeriod = db.studentAttendance.some(a => a.studentId === s.id && a.date.startsWith(prefix));
       } else {
         matchPeriod = db.studentAttendance.some(a => a.studentId === s.id && a.date.startsWith(String(dirFilterYear)));
       }
       return matchSearch && matchSession && matchTutor && matchLevel && matchClass && matchPeriod;
     }));
  }, [db.students, db.studentAttendance, debouncedSearchTerm, filterSession, user, filterLevel, filterClass, dirFilterMonth, dirFilterYear]);

  const filteredTutors = useMemo(() => {
     return db.tutors.filter((t) => {
       const matchSearch = (t.name || '').toLowerCase().includes((debouncedSearchTerm || '').toLowerCase());
       const matchSession = filterSession ? parseSessions(t.teachingSession).includes(filterSession) : true;
       // Fix 2: Month/Year filter for tutors — check tutorAttendance in selected period
       let matchPeriod = true;
       if (dirFilterMonth !== 'All') {
         const prefix = `${dirFilterYear}-${String(dirFilterMonth).padStart(2, '0')}`;
         matchPeriod = db.tutorAttendance.some(a => a.tutorId === t.id && a.date.startsWith(prefix));
       } else {
         matchPeriod = db.tutorAttendance.some(a => a.tutorId === t.id && a.date.startsWith(String(dirFilterYear)));
       }
       return matchSearch && matchSession && matchPeriod;
     });
  }, [db.tutors, db.tutorAttendance, debouncedSearchTerm, filterSession, dirFilterMonth, dirFilterYear]);

  const activeData = dirType === 'student' ? filteredStudents : filteredTutors;
  
  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? activeData.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(activeData.length / (rowsNum || 1));
  const paginatedData = isAll ? activeData : activeData.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;
  // --------------------------------

  const renderDirectory = () => {
    return (
      <div className="space-y-6 animation-fade-in print-hidden">
        <div>
          <div><h2 className="text-2xl font-bold text-white mb-1">
            History & Reports Directory
          </h2><p className="text-gray-400 text-sm">View and export historical attendance, assessment, and payment records.</p></div>
        </div>
        <div className="flex gap-2 p-1 bg-[#151B26] rounded-lg w-max border border-gray-800">
          <button onClick={() => { setDirType('student'); setFilterSession(''); setFilterLevel(''); setFilterClass(''); }} className={`px-6 py-2 rounded-md text-sm font-medium ${dirType === 'student' ? 'bg-[#0B0F19] text-[#00D4FF]' : 'text-gray-400'}`}>Student History</button>
          {user?.role === 'admin' && (
             <button onClick={() => { setDirType('tutor'); setFilterSession(''); setFilterLevel(''); setFilterClass(''); }} className={`px-6 py-2 rounded-md text-sm font-medium ${dirType === 'tutor' ? 'bg-[#0B0F19] text-[#00D4FF]' : 'text-gray-400'}`}>Tutor History</button>
          )}
        </div>
        <Card className="p-0 overflow-hidden flex flex-col">
          <div className="p-4 border-b border-gray-800 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 bg-[#0A0E17]">
            {/* Fix 2: Month & Year filter */}
            <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={dirFilterMonth} onChange={(e) => setDirFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
              <option value="All">All Months</option>
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input type="number" className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={dirFilterYear} onChange={(e) => setDirFilterYear(Number(e.target.value))} placeholder="Year" />
            <input type="text" placeholder={`Search ${dirType}s...`} className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
            
            {dirType === 'student' && (
               <>
                 <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterLevel} onChange={(e) => { setFilterLevel(e.target.value); setFilterClass(''); }}>
                   <option value="">All Levels</option>
                   {LEVELS.map((l) => <option key={l} value={l}>{l}</option>)}
                 </select>
                 <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterClass} onChange={(e) => setFilterClass(e.target.value)} disabled={!filterLevel}>
                   <option value="">All Classes</option>
                   {filterLevel && CLASS_MAPPING[filterLevel]?.map((c) => <option key={c} value={c}>{c}</option>)}
                 </select>
               </>
            )}

            <select className="w-full bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterSession} onChange={(e) => setFilterSession(e.target.value)}>
              {user?.role === 'tutor' ? (
                <>
                  <option value="">All My Sessions</option>
                  {parseSessions(user.teachingSession).map(s => <option key={s} value={s}>{s}</option>)}
                </>
              ) : (
                <>
                  <option value="">All Sessions</option>
                  {SESSIONS.map((l) => <option key={l} value={l}>{l}</option>)}
                </>
              )}
            </select>
          </div>
          {/* ── MOBILE CARD LIST (< sm) ── */}
          <div className="sm:hidden divide-y divide-gray-800">
            {paginatedData.length === 0 ? (
              <EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." />
            ) : dirType === 'student' ? paginatedData.map((s: any, index) => (
              <div key={s.id} className="p-4 flex items-center justify-between gap-3 hover:bg-[#0B0F19]/60 transition-colors">
                <div className="min-w-0">
                  <div className="flex flex-wrap items-center gap-1 mb-1">
                    <span className="font-bold text-white text-sm truncate">{s.name}</span>
                    <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} />
                  </div>
                  <p className="font-mono text-[11px] text-gray-500">{s.id} · {s.class} · <Badge status={s.status} /></p>
                </div>
                <Button className="shrink-0 bg-yellow-500 text-yellow-900 hover:bg-yellow-400 font-bold text-xs px-3 py-1.5 h-auto" icon={Eye} onClick={() => openProfile(s.id, 'student')}>View</Button>
              </div>
            )) : paginatedData.map((t: any, index) => (
              <div key={t.id} className="p-4 flex items-center justify-between gap-3 hover:bg-[#0B0F19]/60 transition-colors">
                <div className="min-w-0">
                  <p className="font-bold text-white text-sm mb-1">{t.name}</p>
                  <p className="font-mono text-[11px] text-gray-500">{t.id} · <Badge status={t.status} /></p>
                </div>
                <Button className="shrink-0 bg-yellow-500 text-yellow-900 hover:bg-yellow-400 font-bold text-xs px-3 py-1.5 h-auto" icon={Eye} onClick={() => openProfile(t.id, 'tutor')}>View</Button>
              </div>
            ))}
          </div>

          {/* ── DESKTOP TABLE (≥ sm) ── */}
          <div className="hidden sm:block overflow-x-auto">
            {dirType === 'student' ? (
              <table className="w-full text-left text-sm">
                <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
                  <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">Student ID</th><th className="p-4 text-center">Name</th><th className="p-4 text-center">Class</th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Actions</th></tr>
                </thead>
                <tbody className="divide-y divide-gray-800">
                  {paginatedData.map((s: any, index) => (
                    <tr key={s.id} className="hover:bg-[#0B0F19]">
                      <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                      <td className="p-4 text-center font-mono text-gray-400">{s.id}</td>
                      <td className="p-4 text-center text-white"><div className="flex items-center justify-center">{s.name} <NewBadge isNew={s.enrollmentStatus} billingStartMonth={s.billingStartMonth || null} /></div></td>
                      <td className="p-4 text-center text-gray-300">{s.class}</td>
                      <td className="p-4 text-center"><Badge status={s.status} /></td>
                      <td className="p-4 flex justify-center"><Button className="bg-yellow-500 text-yellow-900 hover:bg-yellow-400 font-bold shadow-md hover:shadow-lg transition-all" icon={Eye} onClick={() => openProfile(s.id, 'student')}>VIEW HISTORY & REPORTS</Button></td>
                    </tr>
                  ))}
                  {paginatedData.length === 0 && <tr><td colSpan={6}><EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." /></td></tr>}
                </tbody>
              </table>
            ) : (
              <table className="w-full text-left text-sm">
                <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
                  <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">Tutor ID</th><th className="p-4 text-center">Name</th><th className="p-4 text-center">Session</th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Actions</th></tr>
                </thead>
                <tbody className="divide-y divide-gray-800">
                  {paginatedData.map((t: any, index) => (
                    <tr key={t.id} className="hover:bg-[#0B0F19]">
                      <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                      <td className="p-4 text-center font-mono text-gray-400">{t.id}</td>
                      <td className="p-4 text-center text-white">{t.name}</td>
                      <td className="p-4 text-center text-gray-300">{t.teachingSession}</td>
                      <td className="p-4 text-center"><Badge status={t.status} /></td>
                      <td className="p-4 flex justify-center"><Button className="bg-yellow-500 text-yellow-900 hover:bg-yellow-400 font-bold shadow-md hover:shadow-lg transition-all" icon={Eye} onClick={() => openProfile(t.id, 'tutor')}>VIEW HISTORY & REPORTS</Button></td>
                    </tr>
                  ))}
                  {paginatedData.length === 0 && <tr><td colSpan={6}><EmptyState icon={Search} title="No records found" description="Try adjusting your search or filter criteria." /></td></tr>}
                </tbody>
              </table>
            )}
          </div>

          {/* Pagination Footer */}
          <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
             <div className="flex items-center gap-2">
               <span>Show</span>
               <select 
                 value={rowsPerPage} 
                 onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} 
                 className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer"
               >
                 <option value={10}>10</option>
                 <option value={20}>20</option>
                 <option value={30}>30</option>
                 <option value="All">All</option>
               </select>
               <span>entries {activeData.length > 0 && `(Total: ${activeData.length})`}</span>
             </div>
             
             {!isAll && totalPages > 1 && (
               <div className="flex items-center gap-2">
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                 <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
               </div>
             )}
          </div>
        </Card>
      </div>
    );
  };

  const renderStudentProfile = () => {
    if (!student) return null;
    
    const startMonthStr = startDate.substring(0, 7);
    const endMonthStr = endDate.substring(0, 7);
    
    // Fix 4: Only count attendance from student's current effective session (respects override)
    const att = db.studentAttendance.filter((a) => a.studentId === student.id && a.date >= startDate && a.date <= endDate && isAttendanceValidForStudent(a, student));
    const presentCount = att.filter((a) => a.status === 'Present').length;
    const sickCount = att.filter((a) => a.status === 'Sick').length;
    const excusedCount = att.filter((a) => a.status === 'Excused').length;
    const absentCount = att.filter((a) => a.status === 'Absent').length;
    const attRate = att.length ? Math.round((presentCount / att.length) * 100) : 0;
    
    const assessments = db.assessments.filter((a) => {
       if (a.studentId !== student.id) return false;
       const assessDateStr = `${a.year}-${String(a.month).padStart(2, '0')}`;
       return assessDateStr >= startMonthStr && assessDateStr <= endMonthStr;
    });
    const avgScore = assessments.length ? Math.round(assessments.reduce((sum, a) => sum + a.average, 0) / assessments.length) : 0;
    
    const payments = db.payments.filter((p) => p.studentId === student.id && p.date >= startDate && p.date <= endDate);
    const _studentSessionForReport = getStudentSession(student);
    const journals = db.journals.filter((j) => {
      if (j.date < startDate || j.date > endDate) return false;
      const jg = (j.sessionGroup || '').toLowerCase().trim();
      const sg = (_studentSessionForReport || '').toLowerCase().trim();
      return jg === sg || jg.includes(sg) || sg.includes(jg);
    });

    const studentSession = getStudentSession(student);
    const reportSubjects = studentSession === SESSIONS[0]
      ? ['Reading', 'Writing', 'Math', 'English']
      : ['Speaking', 'Writing', 'Reading', 'Listening'];

    const isKindergarten = student.level === 'Kindergarten' || ['PAUD', 'TK A', 'TK B'].includes(student.class);
    // Fix #9: sort descending by year+month agar [0] selalu merupakan record terbaru
    const sortedAssessments = [...assessments].sort((a, b) => {
      if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
      return Number(b.month) - Number(a.month);
    });
    const latestAss = sortedAssessments.length > 0 ? sortedAssessments[0] : null;
    const scores = latestAss && latestAss.scores ? latestAss.scores : {};
    const getScore = (subject) => Number(scores[subject]) || 0;

    const skillProgress = isKindergarten ? [
       { label: 'Reading', value: getScore('Reading') },
       { label: 'Writing', value: getScore('Writing') },
       { label: 'Math', value: getScore('Math') },
       { label: 'English', value: getScore('English') }
    ] : [
       { label: 'Speaking', value: getScore('Speaking') },
       { label: 'Writing', value: getScore('Writing') },
       { label: 'Reading', value: getScore('Reading') },
       { label: 'Listening', value: getScore('Listening') }
    ];

    const hasPaidCurrentMonth = payments.some(p => p.status === 'Paid');
    const paymentStatusBadge = hasPaidCurrentMonth ? <span className="text-green-600 font-bold">Up to Date</span> : <span className="text-red-600 font-bold">Outstanding</span>;
    const totalPaidSum = payments.reduce((sum, p) => sum + Number(p.amount), 0);
    const latestTutor = journals.length > 0 ? journals[journals.length-1].tutorName : "System";

    const autoGeneratedComment = generateAutoComment(student, attRate, avgScore, assessments);
    const finalCommentDisplay = currentTeacherComment || autoGeneratedComment;

    const handleGenerateCommentBtn = () => {
      setCurrentTeacherComment(autoGeneratedComment);
      showToast('Comment generated automatically');
    };

    const handleSaveCommentBtn = () => {
      setDb(p => ({
        ...p,
        // BUGFIX (BUG 8): updatedAt wajib agar LWW tidak menimpa comment guru dengan versi cloud lama
        students: p.students.map(s => s.id === student.id ? {...s, teacherComment: currentTeacherComment, updatedAt: getSyncTimestamp()} : s)
      }));
      showToast('Teacher comment saved to profile');
    };

    const handlePrintStudentReport = () => {
      const originalTitle = document.title;
      const safeName = (student?.name || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_');
      document.title = `${safeName}_report`;
      window.print();
      setTimeout(() => { document.title = originalTitle; }, 1000);
    };

    return (
      <div className="w-full animation-fade-in relative text-slate-900 bg-white rounded-xl shadow-2xl print:shadow-none print:w-full font-sans max-w-5xl mx-auto" id="report-print">
        
        {/* Control Bar & Editor (Print Hidden) */}
        <div className="p-8 pb-0 print:hidden">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 bg-[#0B0F19] p-4 rounded-lg text-white gap-4">
            <Button variant="ghost" onClick={() => setView('directory')} icon={ArrowLeft}>Back</Button>

            {/* Month / Year picker — drives startDate & endDate automatically */}
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs text-gray-400 font-medium whitespace-nowrap">Report Period:</span>
              <select
                className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-sm text-white focus:border-[#00D4FF] focus:outline-none"
                value={reportPickerMonth}
                onChange={e => setReportPickerMonth(Number(e.target.value))}
              >
                {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
              </select>
              <input
                type="number"
                className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-sm text-white focus:border-[#00D4FF] focus:outline-none w-24"
                value={reportPickerYear}
                onChange={e => setReportPickerYear(Number(e.target.value))}
              />
              {/* Show resolved date range as a subtle hint */}
              <span className="text-[11px] text-gray-600 whitespace-nowrap hidden sm:inline">
                ({startDate} → {endDate})
              </span>
            </div>

            <div className="flex items-center gap-2 flex-wrap justify-end">
              <Button onClick={handlePrintStudentReport} icon={Printer}>Print Premium PDF</Button>
            </div>
          </div>

          <div className="mb-8 bg-[#151B26] p-5 rounded-xl border border-gray-800 shadow-xl">
              <h3 className="text-white font-bold mb-3 flex items-center gap-2"><MessageSquare size={18} className="text-[#00D4FF]" /> Report Teacher Comments Editor</h3>
              <p className="text-gray-400 text-xs mb-3">Auto-generate academic comments based on student performance, edit as needed, and save to profile.</p>
              <textarea
                  className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg p-3 text-white text-sm h-28 focus:border-[#00D4FF] transition-colors mb-4"
                  value={currentTeacherComment}
                  onChange={(e) => setCurrentTeacherComment(e.target.value)}
                  placeholder="Write comments or click generate..."
              />
              <div className="flex justify-end gap-3">
                  <Button onClick={handleGenerateCommentBtn} variant="secondary" icon={RefreshCw}>
                     {currentTeacherComment ? 'Regenerate Comment' : 'Generate Comment'}
                  </Button>
                  <Button onClick={handleSaveCommentBtn} icon={CheckCircle2} className="bg-green-600 hover:bg-green-500 border-none text-white shadow-none">Save Comment</Button>
              </div>
          </div>
        </div>

        {/* Actual Print Area */}
        <div className="w-full relative bg-white p-8 border-2 border-[#1A56DB] print:p-0 print:border-2 print:border-[#1A56DB] font-sans max-w-4xl mx-auto text-[11px] text-slate-900">
          {/* Subtle Watermark */}
          <div className="fixed inset-0 flex items-center justify-center opacity-[0.02] z-0 pointer-events-none hidden print:flex">
            <img src={LOGO_URL} className="w-[400px] h-auto grayscale" alt="watermark" />
          </div>

          <div className="relative z-10">
            {/* 1. PREMIUM HEADER */}
            <div className="flex justify-between items-center border-b border-slate-200 pb-4 mb-5 print:mt-4">
              <div className="flex items-center gap-5">
                <img src={LOGO_URL} className="h-16 object-contain" alt="Logo" />
              </div>
              <div className="text-right">
                <h1 className="text-[26px] font-bold text-slate-900 tracking-tight uppercase leading-none">Monthly Academic Progress Report</h1>
                <h2 className="text-[11px] font-medium text-slate-500 mt-2 uppercase tracking-[0.8px]">English Club Gresik • Report Period: {safeDateDisplay(startDate, 'en-GB', { day: 'numeric', month: 'short', year: 'numeric' })} - {safeDateDisplay(endDate, 'en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</h2>
              </div>
            </div>

            {/* 2. STUDENT PROFILE & DASHBOARD SUMMARY CARDS */}
            <div className="grid grid-cols-1 md:grid-cols-12 print:grid-cols-12 gap-3 mb-4 break-inside-avoid">
              {/* Profile Card */}
              <div className="md:col-span-5 print:col-span-5 bg-[#F8FAFC] border border-slate-200 rounded-[10px] p-4">
                <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Student Profile</p>
                <h4 className="text-lg font-bold text-slate-900 mb-0.5">{student.name}</h4>
                <p className="text-[11px] font-mono text-slate-500 mb-3">{student.id}</p>
                <div className="flex gap-6 text-[11px]">
                  <div><span className="text-slate-400 font-medium text-[11px] uppercase tracking-[0.5px] block mb-0.5">Class Level</span><span className="font-semibold text-slate-900">{student.level} - {student.class}</span></div>
                  <div><span className="text-slate-400 font-medium text-[11px] uppercase tracking-[0.5px] block mb-0.5">Session</span><span className="font-semibold text-slate-900">{studentSession}</span></div>
                </div>
              </div>

              {/* Skill Radar Chart */}
              <div className="md:col-span-3 print:col-span-3 bg-white border border-slate-200 rounded-[10px] p-3 shadow-sm flex flex-col items-center justify-center relative">
                 <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px] absolute top-3 left-3">Skill Analytics</p>
                 <div className="w-full max-w-[110px] aspect-square mt-3">
                     <RadarChart data={skillProgress} theme="light" />
                 </div>
              </div>

              {/* Dashboard Summary Cards */}
              <div className="md:col-span-4 print:col-span-4 flex flex-col gap-2">
                 {/* Premium Attendance Card */}
                 <div className="flex-1 bg-white border border-slate-200 rounded-[10px] p-2.5 shadow-sm flex flex-col justify-between">
                    <div className="flex justify-between items-center mb-1">
                       <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px]">Attendance</p>
                       <p className="text-xl font-black text-slate-900 leading-none">{(att.length > 0) ? `${attRate}%` : '—'}</p>
                    </div>
                    {att.length > 0 && (
                       <div className="bg-[#F8FAFC] border border-slate-100 rounded-lg p-1.5 grid grid-cols-2 gap-x-2 gap-y-0.5 text-[8px]">
                          <div className="flex justify-between"><span className="text-slate-500">Present</span><span className="font-bold">{presentCount}</span></div>
                          <div className="flex justify-between"><span className="text-slate-500">Sick</span><span className="font-bold">{sickCount}</span></div>
                          <div className="flex justify-between"><span className="text-slate-500">Excused</span><span className="font-bold">{excusedCount}</span></div>
                          <div className="flex justify-between"><span className="text-slate-500">Absent</span><span className="font-bold">{absentCount}</span></div>
                       </div>
                    )}
                 </div>

                 {/* Premium Avg Score Card */}
                 <div className="flex-1 bg-white border border-slate-200 rounded-[10px] p-2.5 shadow-sm flex flex-col justify-between">
                    <div className="flex justify-between items-center mb-1">
                       <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px]">Avg Score</p>
                       <p className="text-xl font-black text-slate-900 leading-none">{assessments.length > 0 ? avgScore : '—'}</p>
                    </div>
                    {assessments.length > 0 && (
                       <div className="bg-[#F8FAFC] border border-slate-100 rounded-lg p-1.5 flex items-center justify-between h-[30px]">
                          <span className="text-[11px] text-slate-500 font-medium uppercase tracking-wide">Final Grade</span>
                          {/* Fix #9: pakai sortedAssessments[0] — sudah diurutkan terbaru */}
                          <span className="font-black text-base text-blue-600 leading-none">{sortedAssessments[0]?.grade || '—'}</span>
                       </div>
                    )}
                 </div>
              </div>
            </div>

            {/* 3. ACADEMIC ADVISOR COMMENTS */}
            <div className="mb-5 break-inside-avoid">
               <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Academic Advisor Comments</p>
               <div className="bg-[#F8FAFC] border-l-4 border-blue-600 p-[16px] rounded-[10px]">
                  <p className={`text-[11px] leading-relaxed text-slate-800 ${(!assessments.length) ? 'italic text-slate-500' : 'font-medium'}`}>
                     {assessments.length ? `"${finalCommentDisplay}"` : "No comments recorded for this period."}
                  </p>
               </div>
            </div>

            {/* 4. ASSESSMENT TABLE */}
            <div className="mb-5 break-inside-avoid">
              <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Monthly Assessment Grades</p>
              {(assessments.length > 0) ? (
                  <div className="w-full border border-slate-200 rounded-[10px] overflow-hidden">
                      <table className="w-full text-[11px] text-left">
                        <thead className="bg-[#EFF6FF] text-[#1E3A8A]">
                          <tr>
                            <th className="py-2.5 px-4 font-semibold w-24">Period</th>
                            {reportSubjects.map(sub => (
                              <th key={sub} className="py-2.5 px-4 text-center font-semibold">{sub}</th>
                            ))}
                            <th className="py-2.5 px-4 text-center font-semibold w-20">Average</th>
                            <th className="py-2.5 px-4 text-center font-semibold w-16">Grade</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {assessments.map((a, idx) => (
                            <tr key={a.id} className="bg-white even:bg-[#FAFAFA]">
                              <td className="py-2.5 px-4 font-semibold text-slate-900">{MONTHS[parseInt(a.month)-1].substring(0,3)} '{String(a.year).slice(2)}</td>
                              {reportSubjects.map(sub => (
                                <td key={sub} className="py-2.5 px-4 text-center text-slate-700">{a.scores?.[sub] || '—'}</td>
                              ))}
                              <td className="py-2.5 px-4 text-center font-bold text-slate-900">{a.average || '—'}</td>
                              <td className="py-2.5 px-4 text-center font-bold text-blue-600">{a.grade || '—'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                  </div>
              ) : (
                  <div className="bg-[#F8FAFC] border border-slate-200 rounded-[10px] py-3 px-4 flex items-center gap-2 text-slate-500 w-max">
                      <FileText size={14} className="text-slate-400" />
                      <span className="italic text-[11px]">No assessments recorded for this period.</span>
                  </div>
              )}
            </div>

            {/* 5. SESSION DETAIL & ATTENDANCE LOG */}
            <div className="mb-6 break-inside-avoid">
              <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Session Detail & Attendance Log</p>
              {(() => {
                // FIX: Anchor ke attendance siswa saja (bukan union dengan journal).
                // Journal diisi per-sesi oleh tutor dan tidak punya studentId,
                // sehingga union date menyebabkan baris "hantu" — topic tanpa status
                // atau status tanpa topic. Dengan anchor ke att, setiap baris dijamin
                // punya status kehadiran, lalu topic di-lookup dari journal yang tanggalnya sama.
                // Jika tutor belum isi jurnal untuk tanggal itu, tampil "No lesson record entered".
                const sortedAtt = [...att].sort((a, b) =>
                  new Date(a.date).getTime() - new Date(b.date).getTime()
                );

                const combinedSessions = sortedAtt.map((attEntry, idx) => {
                    const journalEntry = journals.find(j => j.date === attEntry.date && j.scheduleId === attEntry.scheduleId)
                      || journals.find(j => j.date === attEntry.date);
                    const status = attEntry.status || "—";
                    const sessionNum = idx + 1;
                    // Calculate streak: how many consecutive Present sessions up to this point
                    const streak = (() => {
                      let s = 0;
                      for (let i = idx; i >= 0; i--) {
                        if (sortedAtt[i].status === 'Present') s++; else break;
                      }
                      return s;
                    })();

                    const autoNote = (() => {
                      if (status === 'Present') {
                        if (sessionNum === 1) return "First session — student attended and engaged well.";
                        if (streak >= 5) return `Excellent consistency — ${streak} sessions in a row with full attendance.`;
                        if (streak >= 3) return `Good streak — attended ${streak} consecutive sessions.`;
                        const notes = [
                          "Attended and followed lesson activities.",
                          "Active participation noted during the session.",
                          "Session completed; student was present and attentive.",
                          "Student followed all learning materials for this session.",
                          "Good attendance — lesson was conducted as scheduled.",
                        ];
                        return notes[sessionNum % notes.length];
                      }
                      if (status === 'Sick') return "Student was absent due to illness. Material may need a brief review next session.";
                      if (status === 'Excused') return "Absence was excused. Student is advised to review session material independently.";
                      if (status === 'Absent') return "Student was absent without notice. Follow-up with parent/guardian recommended.";
                      return "Attendance status unrecorded for this session.";
                    })();

                    return {
                        date: attEntry.date,
                        topic: journalEntry?.topic || "No lesson record entered",
                        status,
                        note: (journalEntry?.notes && journalEntry.notes.trim() !== "") ? journalEntry.notes : autoNote
                    };
                });

                return combinedSessions.length > 0 ? (
                  <div className="w-full border border-slate-200 rounded-[10px] overflow-hidden">
                      <table className="w-full text-[11px] text-left">
                        <thead className="bg-[#EFF6FF] text-[#1E3A8A]">
                          <tr>
                            <th className="py-2.5 px-4 font-semibold w-10 text-center">No</th>
                            <th className="py-2.5 px-4 font-semibold w-24">Date</th>
                            <th className="py-2.5 px-4 font-semibold">Material / Topic</th>
                            <th className="py-2.5 px-4 font-semibold text-center w-24">Status</th>
                            <th className="py-2.5 px-4 font-semibold text-right hidden sm:table-cell print:table-cell">Note</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {combinedSessions.map((session, index) => (
                            <tr key={session.date} className="bg-white even:bg-[#FAFAFA]">
                              <td className="py-2.5 px-4 text-center text-slate-500">{index + 1}</td>
                              <td className="py-2.5 px-4 whitespace-nowrap font-semibold text-slate-900">{new Date(session.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}</td>
                              <td className={`py-2.5 px-4 ${session.topic === 'No lesson record entered' ? 'text-slate-400 italic' : 'text-slate-700'}`}>{session.topic}</td>
                              <td className="py-2.5 px-4 text-center">
                                 <span className="font-bold text-[11px] tracking-wider uppercase">
                                    {session.status !== '—' ? session.status : '—'}
                                 </span>
                              </td>
                              <td className="py-2.5 px-4 text-right text-slate-400 italic text-[11px] hidden sm:table-cell print:table-cell truncate max-w-[180px]">{session.note}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                  </div>
                ) : (
                  <div className="bg-[#F8FAFC] border border-slate-200 rounded-[10px] py-3 px-4 flex items-center gap-2 text-slate-500 w-max">
                      <CalendarIcon size={14} className="text-slate-400" />
                      <span className="italic text-[11px]">No attendance records available for this period.</span>
                  </div>
                );
              })()}
            </div>

            {/* 6. SIGNATURE SECTION */}
            <div className="flex justify-between items-end break-inside-avoid text-[11px] signature-section mb-6">
               <div></div>
               <div className="text-right w-48">
                 <p className="text-slate-500 mb-8 font-medium">Gresik, {new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
                 <div className="border-t border-slate-300 pt-1.5 inline-block w-full">
                     <p className="font-bold text-slate-900 tracking-wide">Akhmad Akmal Rifqi</p>
                     <p className="text-slate-500">Academic Advisor</p>
                 </div>
                 <p className="text-[11px] font-medium text-slate-400 mt-1 flex items-center justify-end gap-1"><CheckCircle2 size={10}/> Verified by ECG Academic Suite</p>
               </div>
            </div>

            {/* 7. PREMIUM FOOTER */}
            <div className="border-t border-slate-200 pt-3 text-[11px] text-slate-400 flex justify-between items-center hidden print:flex uppercase tracking-wide">
               <span>English Club Gresik Premium Report</span>
               <span>Generated automatically</span>
               <span>www.englishclub.my.id</span>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderTutorProfile = () => {
    const tutor = db.tutors.find((t) => t.id === selectedId);
    if (!tutor) return null;
    
    const startMonthStr = startDate.substring(0, 7);
    const endMonthStr = endDate.substring(0, 7);

    const att = db.tutorAttendance.filter((a) => a.tutorId === tutor.id && a.date >= startDate && a.date <= endDate);
    const payrolls = db.payroll.filter((p) => {
        if (p.tutorId !== tutor.id) return false;
        const payDateStr = `${p.year}-${String(p.month).padStart(2, '0')}`;
        return payDateStr >= startMonthStr && payDateStr <= endMonthStr;
    });
    const journals = db.journals.filter((j) => j.tutorName === tutor.name && j.date >= startDate && j.date <= endDate);
    const assessments = db.assessments.filter((a) => {
        if (!parseSessions(tutor.teachingSession).includes(a.sessionGroup)) return false;
        const assessDateStr = `${a.year}-${String(a.month).padStart(2, '0')}`;
        return assessDateStr >= startMonthStr && assessDateStr <= endMonthStr;
    });

    const handlePrintTutorReport = () => {
       const originalTitle = document.title;
       const safeName = (tutor?.name || 'tutor').replace(/[\s/\\?%*:|"<>-]/g, '_');
       document.title = `${safeName}_report`;
       window.print();
       setTimeout(() => { document.title = originalTitle; }, 1000);
    };

    return (
      <div className="w-full animation-fade-in relative text-black bg-white rounded-xl shadow-2xl p-8 print:p-0 print:shadow-none print:bg-transparent print:w-full" id="report-print">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center print-hidden mb-6 bg-[#0B0F19] p-4 rounded-lg text-white gap-4">
          <Button variant="ghost" onClick={() => setView('directory')} icon={ArrowLeft}>Back</Button>
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs text-gray-400 font-medium whitespace-nowrap">Report Period:</span>
            <select
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-sm text-white focus:border-[#00D4FF] focus:outline-none"
              value={reportPickerMonth}
              onChange={e => setReportPickerMonth(Number(e.target.value))}
            >
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input
              type="number"
              className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-1.5 text-sm text-white focus:border-[#00D4FF] focus:outline-none w-24"
              value={reportPickerYear}
              onChange={e => setReportPickerYear(Number(e.target.value))}
            />
            <span className="text-[11px] text-gray-600 whitespace-nowrap hidden sm:inline">
              ({startDate} → {endDate})
            </span>
          </div>
          <Button onClick={handlePrintTutorReport} icon={Printer}>Print / Export PDF</Button>
        </div>

        <table className="w-full relative z-10 text-gray-900 print-table">
          <thead className="table-header-group">
            <tr>
              <td>
                <div className="flex items-center gap-4 border-b-2 border-blue-900 pb-3 mb-4">
                  <img src={LOGO_URL} className="h-14" style={{ filter: 'brightness(0)' }} alt="Logo" />
                  <div>
                    <h1 className="text-xl font-black text-blue-900 tracking-widest uppercase">English Club Gresik</h1>
                    <h2 className="text-base font-bold text-gray-800">Tutor Academic & Performance Report</h2>
                    <p className="text-[11px] text-gray-600">Perumahan Taman Anggrek Blok AB 05, Kedanyang, Kebomas, Gresik | www.englishclub.my.id</p>
                  </div>
                </div>
              </td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <div className="border border-blue-200 bg-blue-50/30 rounded-xl p-4 mb-6 shadow-sm break-inside-avoid">
                  <h3 className="text-xs font-bold text-blue-800 border-b border-gray-200 pb-2 mb-3 uppercase flex items-center gap-2"><User size={14} /> Tutor Profile</h3>
                  <div className="grid grid-cols-2 gap-y-2 text-xs">
                    <div className="flex items-center gap-2"><span className="font-medium w-24 text-gray-500">Name</span> <span className="font-bold">{tutor.name}</span></div>
                    <div className="flex items-center gap-2"><span className="font-medium w-24 text-gray-500">Session</span> <span className="font-bold text-[#1e3a8a]">{parseSessions(tutor.teachingSession).join(', ')}</span></div>
                    <div className="flex items-center gap-2"><span className="font-medium w-24 text-gray-500">Tutor ID</span> <span className="font-bold">{tutor.id}</span></div>
                    <div className="flex items-center gap-2"><span className="font-medium w-24 text-gray-500">Status</span> <span className="font-bold text-blue-700">{tutor.status}</span></div>
                    <div className="flex items-center gap-2"><span className="font-medium w-24 text-gray-500">Generated</span> <span className="font-bold">{new Date().toLocaleString('en-GB')}</span></div>
                  </div>
                </div>

                <div className="mb-6 break-inside-avoid">
                  <h3 className="text-sm font-bold text-blue-900 border-b-2 border-blue-200 mb-2 pb-1 flex items-center gap-2 uppercase">Check-In History</h3>
                  {att.length > 0 ? (
                    <table className="w-full text-[11px] border-collapse border border-gray-300">
                      <thead className="bg-blue-50">
                        <tr>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900">Date</th>
                          <th className="border border-gray-300 p-1.5 text-center font-bold text-blue-900">Time</th>
                          <th className="border border-gray-300 p-1.5 text-center font-bold text-blue-900">Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {att.slice().sort((a,b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map(a => (
                          <tr key={a.id} className="bg-white">
                            <td className="border border-gray-300 p-1.5">{a.date} ({a.day})</td>
                            <td className="border border-gray-300 p-1.5 text-center">{a.time}</td>
                            <td className="border border-gray-300 p-1.5 text-center">{a.status}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : <p className="text-xs text-gray-500 italic">No data available.</p>}
                </div>

                <div className="mb-6 break-inside-avoid">
                  <h3 className="text-sm font-bold text-blue-900 border-b-2 border-blue-200 mb-2 pb-1 flex items-center gap-2 uppercase">Learning Journals</h3>
                  {journals.length > 0 ? (
                    <table className="w-full text-[11px] border-collapse border border-gray-300">
                      <thead className="bg-blue-50">
                        <tr>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900 w-1/4">Date</th>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900 w-1/4">Session</th>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900 w-1/2">Material</th>
                        </tr>
                      </thead>
                      <tbody>
                        {journals.slice().sort((a,b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map(j => (
                          <tr key={j.id} className="bg-white">
                            <td className="border border-gray-300 p-1.5">{j.date}</td>
                            <td className="border border-gray-300 p-1.5">{j.sessionGroup}</td>
                            <td className="border border-gray-300 p-1.5">{j.topic}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : <p className="text-xs text-gray-500 italic">No data available.</p>}
                </div>

                <div className="mb-6 break-inside-avoid">
                  <h3 className="text-sm font-bold text-blue-900 border-b-2 border-blue-200 mb-2 pb-1 flex items-center gap-2 uppercase">Monthly Assessments Conducted</h3>
                  {assessments.length > 0 ? (
                    <table className="w-full text-[11px] border-collapse border border-gray-300">
                      <thead className="bg-blue-50">
                        <tr>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900">Period</th>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900">Student Name</th>
                          <th className="border border-gray-300 p-1.5 text-center font-bold text-blue-900">Avg Score</th>
                          <th className="border border-gray-300 p-1.5 text-center font-bold text-blue-900">Grade</th>
                        </tr>
                      </thead>
                      <tbody>
                        {assessments.slice().sort((a, b) => {
                           if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
                           return Number(b.month) - Number(a.month);
                        }).map(a => (
                          <tr key={a.id} className="bg-white">
                            <td className="border border-gray-300 p-1.5">{MONTHS[parseInt(a.month)-1]} {a.year}</td>
                            <td className="border border-gray-300 p-1.5">{a.studentName}</td>
                            <td className="border border-gray-300 p-1.5 text-center">{a.average}</td>
                            <td className="border border-gray-300 p-1.5 text-center font-bold text-blue-700">{a.grade}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : <p className="text-xs text-gray-500 italic">No data available.</p>}
                </div>

                <div className="mb-6 break-inside-avoid">
                  <h3 className="text-sm font-bold text-blue-900 border-b-2 border-blue-200 mb-2 pb-1 flex items-center gap-2 uppercase">Payroll Summary</h3>
                  {payrolls.length > 0 ? (
                    <table className="w-full text-[11px] border-collapse border border-gray-300">
                      <thead className="bg-blue-50">
                        <tr>
                          <th className="border border-gray-300 p-1.5 text-left font-bold text-blue-900">Period</th>
                          <th className="border border-gray-300 p-1.5 text-right font-bold text-blue-900">Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        {payrolls.slice().sort((a, b) => {
                           if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
                           return Number(b.month) - Number(a.month);
                        }).map(p => (
                          <tr key={p.id} className="bg-white">
                            <td className="border border-gray-300 p-1.5">{MONTHS[parseInt(p.month)-1]} {p.year}</td>
                            <td className="border border-gray-300 p-1.5 text-right">Rp {Number(p.totalPaid).toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : <p className="text-xs text-gray-500 italic">No data available.</p>}
                </div>

              </td>
            </tr>
          </tbody>
          <tfoot className="table-footer-group">
            <tr>
              <td>
                <div className="pt-4 mt-4 border-t border-gray-300 text-[11px] text-gray-500 flex justify-between print-footer-content font-sans">
                   <span>English Club Gresik – Academic Suite</span>
                   <span>Generated on: {new Date().toLocaleString('en-GB')}</span>
                   <span className="print-page-number"></span>
                </div>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    );
  };

  return (
    <div>
      {view === 'directory' && renderDirectory()}
      {view === 'studentProfile' && renderStudentProfile()}
      {view === 'tutorProfile' && renderTutorProfile()}
    </div>
  );
}

function TutorsModule({ db, setDb, generateId, showToast, softDelete }) {
  const [formData, setFormData] = useState({ id: '', name: '', phone: '', address: '', gender: 'Male', teachingSession: SESSIONS[0], status: 'Active', joinedDate: getTodayDateLocal() });
  const [isAdding, setIsAdding] = useState(false);
  
  // NEW: State for Filters & Pagination
  const [filterSession, setFilterSession] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [filterSession, rowsPerPage]);

  const handleSave = (e) => {
    e.preventDefault();
    if (!formData.teachingSession || parseSessions(formData.teachingSession).length === 0) {
      showToast('Please select at least one teaching session', 'warning');
      return;
    }
    // Simpan langsung dalam format internasional (628...) agar siap dipakai untuk link wa.me tanpa konversi lagi.
    const finalPhone = normalizeWhatsapp(formData.phone);
    const rec = { ...formData, phone: finalPhone, id: formData.id || generateId('TUT', 'tutors'), updatedAt: getSyncTimestamp() }; // FIX BUG 3: timestamp untuk LWW
    setDb(p => ({ ...p, tutors: formData.id ? p.tutors.map(t => t.id === formData.id ? rec : t) : [...p.tutors, rec] }));
    showToast('Tutor saved');
    setIsAdding(false);
  };

  const filteredTutors = useMemo(() => {
    return db.tutors.filter(t => filterSession ? parseSessions(t.teachingSession).includes(filterSession) : true);
  }, [db.tutors, filterSession]);

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredTutors.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredTutors.length / (rowsNum || 1));
  const paginatedData = isAll ? filteredTutors : filteredTutors.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white mb-1">Tutors Directory</h2><p className="text-gray-400 text-sm">Manage tutor profiles and teaching sessions.</p></div>
        <Button onClick={() => { setFormData({ id: '', name: '', phone: '', address: '', gender: 'Male', teachingSession: SESSIONS[0], status: 'Active', joinedDate: getTodayDateLocal() }); setIsAdding(!isAdding); }} icon={Plus}>Add Tutor</Button>
      </div>
      {isAdding && (
        <Card className="border border-[#00D4FF]/20 shadow-[0_0_24px_rgba(0,212,255,0.06)]">
          {/* Form Header */}
          <div className="flex items-center gap-3 mb-6 pb-4 border-b border-gray-800">
            <div className="p-2.5 rounded-xl bg-[#00D4FF]/10 border border-[#00D4FF]/20">
              <UserCog size={20} className="text-[#00D4FF]" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white">{formData.id ? 'Edit Tutor' : 'Add New Tutor'}</h3>
              <p className="text-xs text-gray-500 mt-0.5">Fill in all required fields marked with <span className="text-red-400">*</span></p>
            </div>
          </div>

          <form onSubmit={handleSave}>
            {/* Row 1: Full Name (full width) */}
            <div className="mb-4">
              <Input label="Full Name" value={formData.name} onChange={v => setFormData({...formData, name: v})} required />
            </div>

            {/* Row 2: Phone + Gender */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-1">
              <div>
                <Input label="Phone / WhatsApp" value={formData.phone} onChange={v => setFormData({...formData, phone: v})} placeholder="08xxx or 628xxx" required />
                <p className="text-[11px] text-gray-500 mt-1 mb-3 px-1 leading-tight">Format 08... or 628... — auto-saved as international (628...)</p>
              </div>
              <Input label="Gender" type="select" options={['Male', 'Female']} value={formData.gender} onChange={v => setFormData({...formData, gender: v})} required />
            </div>

            {/* Row 3: Teaching Session + Status */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="block text-sm text-gray-400 mb-2">Teaching Sessions <span className="text-red-400">*</span></label>
                <div className="grid grid-cols-1 gap-2">
                  {SESSIONS.map(s => {
                    const checked = parseSessions(formData.teachingSession).includes(s);
                    return (
                      <label key={s} className={`flex items-center gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${checked ? 'border-[#00D4FF]/40 bg-[#00D4FF]/5' : 'border-gray-700 hover:border-gray-600'}`}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => {
                            const cur = parseSessions(formData.teachingSession);
                            const upd = checked ? cur.filter(x => x !== s) : [...cur, s];
                            setFormData({...formData, teachingSession: upd.join('|')});
                          }}
                          className="w-4 h-4 accent-[#00D4FF] shrink-0"
                        />
                        <span className={`text-sm leading-tight ${checked ? 'text-white font-medium' : 'text-gray-400'}`}>{s}</span>
                      </label>
                    );
                  })}
                </div>
              </div>
              <Input label="Status" type="select" options={['Active', 'Inactive']} value={formData.status} onChange={v => setFormData({...formData, status: v})} required />
            </div>

            {/* Row 4: Address (full width) */}
            <div className="mb-6">
              <Input label="Address (Optional)" value={formData.address} onChange={v => setFormData({...formData, address: v})} placeholder="Enter tutor's address..." />
            </div>

            {/* Actions */}
            <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
              <Button variant="secondary" onClick={() => setIsAdding(false)}>Cancel</Button>
              <Button type="submit" icon={formData.id ? Edit2 : Plus}>{formData.id ? 'Update Tutor' : 'Save Tutor'}</Button>
            </div>
          </form>
        </Card>
      )}
      <Card className="p-0 flex flex-col">
        {/* Filter Row */}
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col md:flex-row items-stretch md:items-center gap-3 flex-wrap">
          <select className="w-full md:w-48 bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all" value={filterSession} onChange={(e) => setFilterSession(e.target.value)}>
             <option value="">All Teaching Sessions</option>
             {SESSIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
              <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4">ID</th><th className="p-4">Name</th><th className="p-4">Session</th><th className="p-4">Phone</th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Actions</th></tr>
            </thead>
            <tbody className="divide-y divide-gray-800">
              {paginatedData.map((t, i) => (
                <tr key={t.id} className="hover:bg-[#0B0F19]">
                  <td className="p-4 text-center text-gray-500">{startIndex + i + 1}</td>
                  <td className="p-4 font-mono text-gray-400">{t.id}</td>
                  <td className="p-4 text-white font-medium">{t.name}</td>
                  <td className="p-4 text-gray-300 text-xs">{parseSessions(t.teachingSession).join(' · ') || '-'}</td>
                  <td className="p-4 text-gray-400">{t.phone}</td>
                  <td className="p-4 text-center"><Badge status={t.status} /></td>
                  <td className="p-4 text-center flex justify-center gap-2">
                    {t.phone && (
                      <a href={`https://wa.me/${normalizeWhatsapp(t.phone)}`} target="_blank" rel="noopener noreferrer" className="text-green-400 p-2.5 hover:bg-green-500/10 rounded-lg transition-colors" title="Chat WhatsApp"><MessageCircle size={18}/></a>
                    )}
                    <button onClick={() => { setFormData(t); setIsAdding(true); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Tutor"><Edit2 size={18}/></button>
                    <button onClick={() => softDelete('tutors', t.id, t.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Tutor"><Trash2 size={18}/></button>
                  </td>
                </tr>
              ))}
              {paginatedData.length === 0 && <tr><td colSpan={7}><EmptyState icon={UserCog} title="No tutors found" description="Add a tutor to get started." /></td></tr>}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer */}
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
           <div className="flex items-center gap-2">
             <span>Show</span>
             <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
               <option value={10}>10</option>
               <option value={20}>20</option>
               <option value={30}>30</option>
               <option value="All">All</option>
             </select>
             <span>entries {filteredTutors.length > 0 && `(Total: ${filteredTutors.length})`}</span>
           </div>
           
           {!isAll && totalPages > 1 && (
             <div className="flex items-center gap-2">
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
               <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
               <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
             </div>
           )}
        </div>
      </Card>
    </div>
  );
}

function JournalsModule({ db, setDb, user, showToast, generateId, softDelete }) {
  const [formData, setFormData] = useState({ id: '', scheduleId: '', date: getTodayDateLocal(), sessionGroup: SESSIONS[0], topic: '', activities: '', followUp: '' });
  const [isAdding, setIsAdding] = useState(false);
  // Fix 8: Persistent month/year filter
  const [month, setMonth] = useLocalStorage<number>('ecg_journals_month', new Date().getMonth() + 1);
  const [year, setYear] = useLocalStorage<number>('ecg_journals_year', new Date().getFullYear());
  
  // State for schedule picker filter (separate from list filter)
  const [pickerMonth, setPickerMonth] = useState(new Date().getMonth() + 1);
  const [pickerYear, setPickerYear] = useState(new Date().getFullYear());

  // NEW: State for Filters & Pagination
  const [filterSession, setFilterSession] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [month, year, filterSession, rowsPerPage]);

  // Schedules that already have a journal entry (by this tutor or any tutor, keyed by scheduleId)
  const filledScheduleIds = useMemo(() => {
    const ids = new Set<string>();
    db.journals.forEach(j => {
      if (j.scheduleId) ids.add(String(j.scheduleId));
    });
    return ids;
  }, [db.journals]);

  // NEW: Filter available schedules from Calendar
  const availableSchedules = useMemo(() => {
     let scheds = [...(db.calendar || [])]
        .filter(c => c.date <= getTodayDateLocal())
        .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
     if (user.role === 'tutor') {
        scheds = scheds.filter(c => c.tutor && c.tutor.split(' & ').includes(user.name));
     }
     return scheds;
  }, [db.calendar, user]);

  const handleScheduleChange = (schedId) => {
     const sched = availableSchedules.find(s => s.id === schedId);
     if (sched) {
        setFormData({...formData, scheduleId: sched.id, date: sched.date, sessionGroup: sched.sessionGroup || sched.name});
     } else {
        setFormData({...formData, scheduleId: '', date: '', sessionGroup: ''});
     }
  };

  const filteredJournals = useMemo(() => {
    return db.journals.filter(j => {
      const matchMonth = j.date.startsWith(`${year}-${String(month).padStart(2, '0')}`);
      const matchSession = filterSession ? j.sessionGroup === filterSession : true;
      const matchTutor = user.role === 'admin' ? true : j.tutorName === user.name;
      return matchMonth && matchSession && matchTutor;
    });
  }, [db.journals, month, year, filterSession, user]);

  const reversedJournals = [...filteredJournals].reverse(); // Show newest first

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? reversedJournals.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(reversedJournals.length / (rowsNum || 1));
  const paginatedData = isAll ? reversedJournals : reversedJournals.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  const handleSave = (e) => {
    e.preventDefault();
    // Guard: prevent creating a NEW journal for a schedule that already has one
    if (!formData.id && formData.scheduleId && filledScheduleIds.has(String(formData.scheduleId))) {
      showToast('Journal for this schedule already exists. Please edit the existing one instead.', 'error');
      return;
    }
    // FIX BUG #4: tambah timestamp untuk LWW yang benar di mergeByIds
    const rec = { ...formData, id: formData.id || generateId('JRN', 'journals'), tutorName: user.name, timestamp: getLocalTimestamp() };
    setDb(p => ({ ...p, journals: formData.id ? p.journals.map(j => j.id === formData.id ? rec : j) : [...p.journals, rec] }));
    showToast(formData.id ? 'Journal updated' : 'Journal saved');
    setIsAdding(false);
  };

  // Session color mapping for visual distinction
  const SESSION_COLORS = {
    [SESSIONS[0]]: { bg: 'bg-pink-500/10', border: 'border-pink-500/30', text: 'text-pink-400', dot: 'bg-pink-400', bar: 'bg-pink-500' },
    [SESSIONS[1]]: { bg: 'bg-blue-500/10', border: 'border-blue-500/30', text: 'text-blue-400', dot: 'bg-blue-400', bar: 'bg-blue-500' },
    [SESSIONS[2]]: { bg: 'bg-emerald-500/10', border: 'border-emerald-500/30', text: 'text-emerald-400', dot: 'bg-emerald-400', bar: 'bg-emerald-500' },
    [SESSIONS[3]]: { bg: 'bg-amber-500/10', border: 'border-amber-500/30', text: 'text-amber-400', dot: 'bg-amber-400', bar: 'bg-amber-500' },
    [SESSIONS[4]]: { bg: 'bg-purple-500/10', border: 'border-purple-500/30', text: 'text-purple-400', dot: 'bg-purple-400', bar: 'bg-purple-500' },
  };
  const getSessionColor = (sg) => SESSION_COLORS[sg] || { bg: 'bg-gray-500/10', border: 'border-gray-500/30', text: 'text-gray-400', dot: 'bg-gray-400', bar: 'bg-gray-500' };

  // Stats for the month
  const sessionStats = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredJournals.forEach(j => { counts[j.sessionGroup] = (counts[j.sessionGroup] || 0) + 1; });
    return counts;
  }, [filteredJournals]);

  // Group journals by date for timeline
  const groupedByDate = useMemo(() => {
    const groups: Record<string, typeof paginatedData> = {};
    paginatedData.forEach(j => {
      if (!groups[j.date]) groups[j.date] = [];
      groups[j.date].push(j);
    });
    return Object.entries(groups).sort(([a], [b]) => b.localeCompare(a));
  }, [paginatedData]);

  return (
    <div className="space-y-5">
      {/* Header */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-white mb-1">Learning Journals</h2>
          <p className="text-gray-400 text-sm">Log topics and activities covered in each session.</p>
        </div>
        <Button
          onClick={() => { setFormData({ id: '', scheduleId: '', date: '', sessionGroup: '', topic: '', activities: '', followUp: '' }); setIsAdding(!isAdding); }}
          icon={Plus}
        >
          Write Journal
        </Button>
      </div>

      {/* Write/Edit Form */}
      {isAdding && (
        <Card className="border border-purple-500/25 shadow-[0_0_32px_rgba(168,85,247,0.08)]">
          <div className="flex items-center gap-3 mb-5 pb-4 border-b border-gray-800">
            <div className="p-2.5 rounded-xl bg-purple-500/10 border border-purple-500/20">
              <BookOpen size={20} className="text-purple-400" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white">{formData.id ? 'Edit Journal' : 'Write New Journal'}</h3>
              <p className="text-xs text-gray-500 mt-0.5">Fields marked <span className="text-red-400">*</span> are required</p>
            </div>
          </div>
          <form onSubmit={handleSave}>
            {!formData.id ? (
              <div className="mb-4">
                <label className="block text-sm text-gray-400 mb-2">Select Class Schedule <span className="text-red-400">*</span></label>
                {/* Step 1: Month / Year filter */}
                <div className="flex gap-2 mb-3">
                  <select
                    className="flex-1 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF] transition-all"
                    value={pickerMonth}
                    onChange={e => { setPickerMonth(Number(e.target.value)); handleScheduleChange(''); }}
                  >
                    {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
                  </select>
                  <input
                    type="number"
                    className="w-24 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF] transition-all"
                    value={pickerYear}
                    onChange={e => { setPickerYear(Number(e.target.value)); handleScheduleChange(''); }}
                  />
                </div>
                {/* Step 2: Schedule cards for chosen month */}
                {(() => {
                  const prefix = `${pickerYear}-${String(pickerMonth).padStart(2, '0')}`;
                  const monthScheds = availableSchedules.filter(c => (c.date || '').startsWith(prefix));
                  if (monthScheds.length === 0) return (
                    <p className="text-center text-gray-600 text-sm py-4 bg-[#0B0F19] rounded-lg border border-gray-800">
                      No schedules found for {MONTHS[pickerMonth - 1]} {pickerYear}
                    </p>
                  );
                  return (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 max-h-52 overflow-y-auto custom-scrollbar pr-1">
                      {monthScheds.map(c => {
                        const isSelected = formData.scheduleId === c.id;
                        const isFilled = filledScheduleIds.has(String(c.id));
                        const col = (() => {
                          const sg = c.sessionGroup || c.name || '';
                          if (sg.includes('PAUD') || sg.includes('TK')) return { border: 'border-pink-500/40', bg: 'bg-pink-500/10', text: 'text-pink-400', sel: 'border-pink-400 bg-pink-500/20' };
                          if (sg.includes('Grade 1') || sg.includes('Grade 2')) return { border: 'border-blue-500/40', bg: 'bg-blue-500/10', text: 'text-blue-400', sel: 'border-blue-400 bg-blue-500/20' };
                          if (sg.includes('Grade 3') || sg.includes('Grade 4')) return { border: 'border-emerald-500/40', bg: 'bg-emerald-500/10', text: 'text-emerald-400', sel: 'border-emerald-400 bg-emerald-500/20' };
                          if (sg.includes('Grade 5') || sg.includes('Grade 6')) return { border: 'border-amber-500/40', bg: 'bg-amber-500/10', text: 'text-amber-400', sel: 'border-amber-400 bg-amber-500/20' };
                          return { border: 'border-purple-500/40', bg: 'bg-purple-500/10', text: 'text-purple-400', sel: 'border-purple-400 bg-purple-500/20' };
                        })();
                        const dateObj = new Date(c.date + 'T00:00:00');
                        const dayShort = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
                        const dayNum = dateObj.getDate();
                        return (
                          <button
                            key={c.id}
                            type="button"
                            disabled={isFilled}
                            onClick={() => !isFilled && handleScheduleChange(c.id)}
                            className={`flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                              isFilled
                                ? 'border-emerald-700/40 bg-emerald-900/10 opacity-60 cursor-not-allowed'
                                : isSelected
                                  ? col.sel + ' border-2'
                                  : 'border-gray-700 bg-[#0B0F19] hover:' + col.bg + ' hover:' + col.border
                            }`}
                          >
                            {/* Date badge */}
                            <div className={`flex-shrink-0 w-10 h-10 rounded-lg flex flex-col items-center justify-center ${isFilled ? 'bg-emerald-900/30' : isSelected ? col.bg : 'bg-[#151B26]'}`}>
                              <span className={`text-[10px] font-bold uppercase ${isFilled ? 'text-emerald-500' : col.text}`}>{dayShort}</span>
                              <span className="text-white text-sm font-bold leading-none">{dayNum}</span>
                            </div>
                            {/* Info */}
                            <div className="flex-1 min-w-0">
                              <p className={`text-xs font-semibold truncate ${isFilled ? 'text-emerald-500' : col.text}`}>{c.sessionGroup || c.name}</p>
                              <p className="text-gray-400 text-[11px] mt-0.5">{c.startTime}{c.tutor ? ` · ${c.tutor}` : ''}</p>
                              {isFilled && <p className="text-emerald-500 text-[10px] font-bold mt-0.5">✓ Journal filled — edit from the list</p>}
                            </div>
                            {isFilled ? <CheckCircle2 size={16} className="text-emerald-500 flex-shrink-0" /> : isSelected ? <CheckCircle2 size={16} className={col.text} /> : null}
                          </button>
                        );
                      })}
                    </div>
                  );
                })()}
                {/* Validation hint */}
                {!formData.scheduleId && (
                  <p className="text-[11px] text-gray-600 mt-2">Pick a schedule card above to continue.</p>
                )}
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                <Input label="Date" type="date" value={formData.date} disabled />
                <Input label="Session" type="text" value={formData.sessionGroup} disabled />
              </div>
            )}
            <div className="mb-4">
              <Input label="Topic / Material" value={formData.topic} onChange={v => setFormData({...formData, topic: v})} required />
            </div>
            <div className="mb-4">
              <label className="block text-sm text-gray-400 mb-1.5">Class Activities <span className="text-red-400">*</span></label>
              <textarea className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg p-3 text-white focus:border-[#00D4FF] focus:outline-none transition-colors text-sm" rows={3} value={formData.activities} onChange={e => setFormData({...formData, activities: e.target.value})} required placeholder="Describe the activities done in class..." />
            </div>
            <div className="mb-6">
              <label className="block text-sm text-gray-400 mb-1.5">Follow Up / Notes</label>
              <textarea className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg p-3 text-white focus:border-[#00D4FF] focus:outline-none transition-colors text-sm" rows={2} value={formData.followUp} onChange={e => setFormData({...formData, followUp: e.target.value})} placeholder="Homework, special notes, or plans for the next session..." />
            </div>
            <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
              <Button variant="secondary" onClick={() => setIsAdding(false)}>Cancel</Button>
              <Button type="submit" icon={formData.id ? Edit2 : Plus}>{formData.id ? 'Update Journal' : 'Save Journal'}</Button>
            </div>
          </form>
        </Card>
      )}

      {/* Filter Bar */}
      <div className="bg-[#151B26] border border-gray-800 rounded-xl p-4 flex flex-col md:flex-row gap-3 items-center">
        <div className="flex gap-2 w-full md:w-auto">
          <select
            className="flex-1 md:w-36 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF] transition-all"
            value={month} onChange={e => setMonth(Number(e.target.value))}
          >
            {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
          </select>
          <input
            type="number"
            className="w-20 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF] transition-all"
            value={year} onChange={e => setYear(Number(e.target.value))}
          />
        </div>
        <select
          className="w-full md:w-56 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF] transition-all"
          value={filterSession} onChange={(e) => setFilterSession(e.target.value)}
        >
          {user?.role === 'tutor' ? (
            <>
              <option value="">All My Sessions</option>
              {parseSessions(user.teachingSession).map(s => <option key={s} value={s}>{s}</option>)}
            </>
          ) : (
            <>
              <option value="">All Sessions</option>
              {SESSIONS.map((s) => <option key={s} value={s}>{s}</option>)}
            </>
          )}
        </select>

        {/* Stats chips */}
        {filteredJournals.length > 0 && (
          <div className="flex flex-wrap gap-2 ml-auto">
            {Object.entries(sessionStats).map(([sg, count]) => {
              const c = getSessionColor(sg);
              const shortLabel = sg.replace(' Session', '').replace('Grade ', 'Gr.');
              return (
                <span key={sg} className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-semibold ${c.bg} ${c.text} border ${c.border}`}>
                  <span className={`w-1.5 h-1.5 rounded-full ${c.dot}`} />
                  {shortLabel} · {count}
                </span>
              );
            })}
          </div>
        )}
      </div>

      {/* Timeline / Journal List */}
      {groupedByDate.length === 0 ? (
        <div className="bg-[#151B26] border border-gray-800 rounded-xl py-16 flex flex-col items-center gap-3">
          <div className="w-14 h-14 rounded-full bg-gray-800 flex items-center justify-center">
            <BookOpen size={24} className="text-gray-600" />
          </div>
          <p className="text-gray-500 font-medium">No journals recorded for this period.</p>
          <p className="text-gray-600 text-sm">Klik "Write Journal" untuk menambahkan</p>
        </div>
      ) : (
        <div className="space-y-6">
          {groupedByDate.map(([date, journals]) => {
            const dateObj = new Date(date + 'T00:00:00');
            const dayName = dateObj.toLocaleDateString('id-ID', { weekday: 'long' });
            const dateStr = dateObj.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
            return (
              <div key={date}>
                {/* Date separator */}
                <div className="flex items-center gap-3 mb-3">
                  <div className="bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-1.5 flex items-center gap-2">
                    <CalendarIcon size={13} className="text-[#00D4FF]" />
                    <span className="text-white text-xs font-bold">{dayName},</span>
                    <span className="text-gray-400 text-xs">{dateStr}</span>
                  </div>
                  <div className="flex-1 h-px bg-gray-800" />
                  <span className="text-gray-600 text-xs">{journals.length} jurnal</span>
                </div>

                {/* Cards for this date */}
                <div className="space-y-3 pl-0">
                  {journals.map(j => {
                    const c = getSessionColor(j.sessionGroup);
                    return (
                      <div key={j.id} className={`group relative bg-[#0F1724] border ${c.border} rounded-xl overflow-hidden transition-all hover:shadow-[0_4px_20px_rgba(0,0,0,0.3)] hover:-translate-y-0.5`}>
                        {/* Left accent bar */}
                        <div className={`absolute left-0 top-0 bottom-0 w-1 ${c.bar}`} />

                        <div className="pl-5 pr-4 py-4">
                          {/* Top row */}
                          <div className="flex items-start justify-between gap-3 mb-3">
                            <div className="flex-1 min-w-0">
                              <h3 className="text-white font-bold text-base leading-snug mb-1.5 pr-16">{j.topic}</h3>
                              <div className="flex flex-wrap items-center gap-2">
                                <span className={`inline-flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-full ${c.bg} ${c.text} border ${c.border}`}>
                                  <span className={`w-1.5 h-1.5 rounded-full ${c.dot}`} />
                                  {j.sessionGroup}
                                </span>
                                <span className="text-gray-600 text-xs flex items-center gap-1">
                                  <User size={10} />
                                  {j.tutorName}
                                </span>
                              </div>
                            </div>
                            {/* Action buttons */}
                            <div className="absolute top-3.5 right-3.5 flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button
                                onClick={() => { setFormData(j); setIsAdding(true); window.scrollTo({ top: 0, behavior: 'smooth' }); }}
                                className="text-blue-400 p-2 hover:bg-blue-500/15 rounded-lg transition-colors"
                                title="Edit Journal"
                              ><Edit2 size={15}/></button>
                              <button
                                onClick={() => softDelete('journals', j.id, 'Journal')}
                                className="text-red-400 p-2 hover:bg-red-500/15 rounded-lg transition-colors"
                                title="Delete Journal"
                              ><Trash2 size={15}/></button>
                            </div>
                          </div>

                          {/* Content */}
                          <div className="space-y-2">
                            <div className="bg-[#0B0F19] rounded-lg px-3 py-2.5">
                              <p className="text-[11px] text-gray-500 uppercase font-semibold tracking-wide mb-1">Activities</p>
                              <p className="text-sm text-gray-300 leading-relaxed">{j.activities}</p>
                            </div>
                            {j.followUp && (
                              <div className="bg-amber-500/5 border border-amber-500/15 rounded-lg px-3 py-2.5">
                                <p className="text-[11px] text-amber-500/80 uppercase font-semibold tracking-wide mb-1">Follow Up / Notes</p>
                                <p className="text-sm text-amber-100/70 leading-relaxed">{j.followUp}</p>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Pagination Footer */}
      <div className="bg-[#151B26] border border-gray-800 rounded-xl p-4 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
        <div className="flex items-center gap-2">
          <span className="text-gray-500">Show</span>
          <select
            value={rowsPerPage}
            onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))}
            className="bg-[#0B0F19] border border-gray-700 rounded-lg px-2.5 py-1.5 text-white text-sm focus:outline-none focus:border-[#00D4FF] cursor-pointer"
          >
            <option value={10}>10</option>
            <option value={20}>20</option>
            <option value={50}>50</option>
            <option value="All">All</option>
          </select>
          <span className="text-gray-500">entries {filteredJournals.length > 0 && <span className="text-gray-300 font-medium">· Total: {filteredJournals.length}</span>}</span>
        </div>
        {!isAll && totalPages > 1 && (
          <div className="flex items-center gap-2">
            <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#0B0F19] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>← Prev</Button>
            <span className="px-3 py-1.5 text-white font-medium text-sm">{currentPage} / {totalPages}</span>
            <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#0B0F19] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next →</Button>
          </div>
        )}
      </div>
    </div>
  );
}

function PayrollModule({ db, setDb, generateId, showToast, handlePrint, handleShareImage, downloadPNG, softDelete, requestConfirm }) {
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [year, setYear] = useState(new Date().getFullYear());
  const [selectedPayroll, setSelectedPayroll] = useState(null);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editFormData, setEditFormData] = useState(null);

  // NEW: State for Pagination
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [month, year, rowsPerPage]);

  const activeTutors = db.tutors.filter(t => t.status === 'Active');
  
  const generatePayroll = () => {
     const newPayrolls = [];
     activeTutors.forEach(tutor => {
        const existing = db.payroll.find(p => p.tutorId === tutor.id && Number(p.month) === Number(month) && String(p.year) === String(year));
        if (existing) return;

        const base = Number(tutor.baseSalary) || 0;
        const monthPrefix = `${year}-${String(month).padStart(2, '0')}`;
        const classesDone = db.calendar.filter(c => c.tutor && c.tutor.split(' & ').includes(tutor.name) && c.date.startsWith(monthPrefix) && c.date <= getTodayDateLocal()).length;
        const teachingTotal = base * classesDone;
        const total = teachingTotal;

        newPayrolls.push({
           id: generateId('PAY', 'payroll'),
           tutorId: tutor.id,
           tutorName: tutor.name,
           month: String(month),
           year: String(year),
           baseSalary: base,
           classesDone,
           teachingBonus: teachingTotal,
           transportAllowance: 0,
           deductions: 0,
           totalPaid: total,
           status: 'Draft',
           date: getTodayDateLocal(),
           updatedAt: getSyncTimestamp(), // FIX BUG 3: timestamp untuk LWW
        });
     });
     
     if (newPayrolls.length > 0) {
        setDb(p => ({ ...p, payroll: [...p.payroll, ...newPayrolls] }));
        showToast('Payroll generated for active tutors.');
     } else {
        showToast('All active tutors already have payroll records for this month.', 'warning');
     }
  };

  const handleSaveEdit = (e) => {
     e.preventDefault();
     if(!editFormData) return;
     setDb(p => ({
        ...p,
        payroll: p.payroll.map(pay => {
           if (pay.id === editFormData.id) {
              const teachingTotal = Number(editFormData.baseSalary) * Number(editFormData.classesDone);
              const total = teachingTotal +
                            Number(editFormData.transportAllowance) +
                            Number(editFormData.additionalBonus || 0) -
                            Number(editFormData.deductions);
              return {
                 ...pay,
                 baseSalary: Number(editFormData.baseSalary),
                 classesDone: Number(editFormData.classesDone),
                 teachingBonus: teachingTotal,
                 transportAllowance: Number(editFormData.transportAllowance),
                 additionalBonus: Number(editFormData.additionalBonus || 0),
                 deductions: Number(editFormData.deductions),
                 totalPaid: total,
                 updatedAt: getSyncTimestamp(), // BUGFIX (BUG 7): LWW wajib tahu versi ini lebih baru
              };
           }
           return pay;
        })
     }));
     setIsEditModalOpen(false);
     setEditFormData(null);
     showToast('Payroll updated successfully');
  };

  const markAsPaid = (id) => {
     requestConfirm('Mark as Paid', 'Are you sure you want to finalize this payroll? It will be marked as Paid.', () => {
        setDb(p => ({
           ...p,
           payroll: p.payroll.map(pay => pay.id === id ? { ...pay, status: 'Paid', date: getTodayDateLocal(), updatedAt: getSyncTimestamp() } : pay) // BUGFIX (BUG 7)
        }));
        showToast('Payroll marked as Paid!');
     });
  };

  const filteredPayroll = db.payroll.filter(p => Number(p.month) === Number(month) && String(p.year) === String(year));

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredPayroll.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredPayroll.length / (rowsNum || 1));
  const paginatedPayroll = isAll ? filteredPayroll : filteredPayroll.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  const localPrintPayroll = () => {
    if (!selectedPayroll) return;
    const originalTitle = document.title;
    const safeName = selectedPayroll.tutorName.replace(/[\s/\\?%*:|"<>-]/g, '_');
    document.title = `${safeName}_payroll_${MONTHS[Number(selectedPayroll.month)-1]}_${selectedPayroll.year}`;
    window.print();
    setTimeout(() => { document.title = originalTitle; }, 1000);
  };

  const handleSharePayrollWA = () => {
    if (!selectedPayroll) return;
    const tutor = db.tutors.find(t => t.id === selectedPayroll.tutorId);
    const phone = tutor?.phone ? normalizeWhatsapp(tutor.phone) : '';
    const monthName = MONTHS[Number(selectedPayroll.month) - 1];
    const msg =
      `📋 *PAYROLL SLIP — ${monthName} ${selectedPayroll.year}*\n` +
      `English Club Gresik\n\n` +
      `👤 *${selectedPayroll.tutorName}*\n` +
      `📅 Periode: ${monthName} ${selectedPayroll.year}\n\n` +
      `📊 *Rincian Gaji:*\n` +
      `• Mengajar: ${selectedPayroll.classesDone} kelas × Rp ${Number(selectedPayroll.baseSalary).toLocaleString('id-ID')} = Rp ${(Number(selectedPayroll.baseSalary) * Number(selectedPayroll.classesDone)).toLocaleString('id-ID')}\n` +
      (Number(selectedPayroll.transportAllowance) > 0 ? `• Reimbursement: + Rp ${Number(selectedPayroll.transportAllowance).toLocaleString('id-ID')}\n` : '') +
      (Number(selectedPayroll.additionalBonus) > 0 ? `• Bonus Tambahan: + Rp ${Number(selectedPayroll.additionalBonus).toLocaleString('id-ID')}\n` : '') +
      (Number(selectedPayroll.deductions) > 0 ? `• Potongan: - Rp ${Number(selectedPayroll.deductions).toLocaleString('id-ID')}\n` : '') +
      `\n💰 *Take Home Pay: Rp ${Number(selectedPayroll.totalPaid).toLocaleString('id-ID')}*\n\n` +
      `Status: ${selectedPayroll.status === 'Paid' ? '✅ Sudah Dibayar' : '⏳ Draft'}\n\n` +
      `_Terima kasih atas dedikasi Anda! — ECG Academic Suite_`;
    if (phone) {
      window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, '_blank');
    } else {
      // No phone stored — open WA without pre-filled number
      window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
      showToast('No phone number saved for this tutor. Opening WhatsApp without a pre-filled contact.', 'warning');
    }
  };

  if (selectedPayroll) return (
    <div className="fixed inset-0 z-[100] bg-slate-50/95 backdrop-blur-md overflow-y-auto print:bg-white print:static print:block print:z-auto custom-scrollbar font-sans text-slate-900">
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[80%] h-[400px] bg-gradient-to-b from-blue-200/40 to-transparent blur-3xl pointer-events-none print:hidden" />
      <div className="w-full max-w-2xl mx-auto mt-6 mb-4 px-4 flex justify-between items-center relative z-10 print:hidden">
        <button className="flex items-center gap-2 text-slate-500 hover:text-slate-800 transition-colors font-medium text-sm bg-white/50 px-4 py-2 rounded-full border border-slate-200/50 shadow-sm" onClick={() => setSelectedPayroll(null)}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="flex gap-2 flex-wrap justify-end">
          <button onClick={localPrintPayroll} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Print PDF"><Printer size={16}/></button>
          <button onClick={() => {
             const safeName = selectedPayroll.tutorName.replace(/[\s/\\?%*:|"<>-]/g, '_');
             downloadPNG('payroll-print', `${safeName}_payroll`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Download PNG"><Download size={16}/></button>
          <button onClick={() => {
             const safeName = selectedPayroll.tutorName.replace(/[\s/\\?%*:|"<>-]/g, '_');
             handleShareImage('payroll-print', `${safeName}_payroll`, `Payroll Slip ${selectedPayroll.tutorName} — ${MONTHS[Number(selectedPayroll.month)-1]} ${selectedPayroll.year}`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Share Image"><Share2 size={16}/></button>
          <button onClick={handleSharePayrollWA} className="flex items-center gap-1.5 px-3 py-2 bg-green-500 hover:bg-green-600 text-white rounded-full shadow-sm text-xs font-bold transition-colors" title="Share via WhatsApp">
            <MessageCircle size={14}/> WhatsApp
          </button>
        </div>
      </div>

      <div id="payroll-print" className="w-full max-w-2xl mx-auto bg-white rounded-none shadow-2xl overflow-hidden relative z-10 mb-12 print:m-0 print:shadow-none print:max-w-full print:rounded-none">
        <div className="bg-[#1A56DB] text-white p-6 sm:p-8 flex justify-between items-start">
          <div>
            <h1 className="text-2xl sm:text-3xl font-black tracking-tight">PAYROLL SLIP</h1>
            <p className="text-blue-200 font-mono mt-1 text-sm">{selectedPayroll.id}</p>
          </div>
          <div className="text-right">
            <p className="text-lg font-bold text-white leading-tight">English Club Gresik</p>
            <p className="text-blue-200 text-xs mt-1">Academic Suite</p>
          </div>
        </div>

        <div className="p-6 sm:p-8 pb-4">
          <div className="text-center mb-8 border-b border-slate-200 pb-8">
            <p className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">
              Salary Period {MONTHS[parseInt(selectedPayroll.month) - 1]} {selectedPayroll.year}
            </p>
            <p className="text-5xl font-black text-slate-900 tracking-tight">
              Rp {Number(selectedPayroll.totalPaid).toLocaleString('id-ID')}
            </p>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">TUTOR INFO</p>
            <div className="bg-slate-50 border border-slate-200 p-5 shadow-sm">
              <p className="text-xl font-bold text-slate-800">{selectedPayroll.tutorName}</p>
            </div>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">EARNINGS & DEDUCTIONS</p>
            <div className="space-y-0 text-sm">
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Teaching Salary ({selectedPayroll.classesDone} Classes x Rp {Number(selectedPayroll.baseSalary).toLocaleString('id-ID')})</span>
                <span className="font-semibold text-slate-800">Rp {(Number(selectedPayroll.baseSalary) * Number(selectedPayroll.classesDone)).toLocaleString('id-ID')}</span>
              </div>
              {Number(selectedPayroll.transportAllowance) > 0 && (
                <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                  <span className="font-medium text-slate-600">Reimbursement</span>
                  <span className="font-semibold text-green-600">+ Rp {Number(selectedPayroll.transportAllowance).toLocaleString('id-ID')}</span>
                </div>
              )}
              {Number(selectedPayroll.additionalBonus) > 0 && (
                 <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                   <span className="font-medium text-slate-600">Additional Bonus</span>
                   <span className="font-semibold text-green-600">+ Rp {Number(selectedPayroll.additionalBonus).toLocaleString('id-ID')}</span>
                 </div>
              )}
              {Number(selectedPayroll.deductions) > 0 && (
                <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                  <span className="font-medium text-slate-600">Deductions</span>
                  <span className="font-semibold text-red-600">- Rp {Number(selectedPayroll.deductions).toLocaleString('id-ID')}</span>
                </div>
              )}
              <div className="flex justify-between items-center py-4 px-2 bg-slate-50 mt-2">
                <span className="font-bold text-slate-800">Net Take Home Pay</span>
                <span className="text-lg font-black text-[#1A56DB]">Rp {Number(selectedPayroll.totalPaid).toLocaleString()}</span>
              </div>
            </div>
          </div>

          <div className="text-center mt-12 mb-4">
            <p className="text-sm font-medium text-slate-600 italic">Thank you for your excellent dedication.</p>
            <p className="text-xs font-bold text-slate-500 mt-1">— Akhmad Akmal Rifqi</p>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="space-y-6">
      {isEditModalOpen && editFormData && (
         <CustomModal isOpen={true} onClose={() => setIsEditModalOpen(false)} title={`Edit Payroll: ${editFormData.tutorName}`}>
           <form onSubmit={handleSaveEdit} className="space-y-4">
              <Input label="Base Salary (Rp)" type="number" value={editFormData.baseSalary || ''} onChange={v => setEditFormData({...editFormData, baseSalary: v})} required />
              <Input label="Classes Taught" type="number" value={editFormData.classesDone || ''} onChange={v => setEditFormData({...editFormData, classesDone: v})} required />
              <Input label="Reimbursement (Rp)" type="number" value={editFormData.transportAllowance || ''} onChange={v => setEditFormData({...editFormData, transportAllowance: v})} />
              <Input label="Additional Bonus (Rp)" type="number" value={editFormData.additionalBonus || ''} onChange={v => setEditFormData({...editFormData, additionalBonus: v})} />
              <Input label="Deductions (Rp)" type="number" value={editFormData.deductions || ''} onChange={v => setEditFormData({...editFormData, deductions: v})} />

              <div className="p-4 bg-blue-500/10 border border-blue-500/30 rounded-xl mt-4">
                 <div className="flex justify-between items-center mb-2">
                    <p className="text-xs text-blue-400 font-bold uppercase tracking-wider">Take Home Pay Preview</p>
                 </div>
                 <div className="space-y-1.5 mb-3 text-sm text-gray-300 font-medium">
                    <div className="flex justify-between"><span>Base Salary:</span> <span>Rp {Number(editFormData.baseSalary).toLocaleString('id-ID')}</span></div>
                    <div className="flex justify-between"><span>Classes Taught:</span> <span>{editFormData.classesDone}</span></div>
                    <div className="flex justify-between text-[#00D4FF]"><span>Teaching Subtotal ({editFormData.classesDone} x Rp {Number(editFormData.baseSalary).toLocaleString('id-ID')}):</span> <span>Rp {(Number(editFormData.baseSalary) * Number(editFormData.classesDone)).toLocaleString('id-ID')}</span></div>
                    {Number(editFormData.transportAllowance) > 0 && <div className="flex justify-between text-green-400"><span>Reimbursement:</span> <span>+ Rp {Number(editFormData.transportAllowance).toLocaleString('id-ID')}</span></div>}
                    {Number(editFormData.additionalBonus) > 0 && <div className="flex justify-between text-green-400"><span>Additional Bonus:</span> <span>+ Rp {Number(editFormData.additionalBonus).toLocaleString('id-ID')}</span></div>}
                    {Number(editFormData.deductions) > 0 && <div className="flex justify-between text-red-400"><span>Deductions:</span> <span>- Rp {Number(editFormData.deductions).toLocaleString('id-ID')}</span></div>}
                 </div>
                 <div className="pt-3 border-t border-blue-500/30">
                    <p className="text-2xl font-black text-white">
                       Rp {(
                          (Number(editFormData.baseSalary) * Number(editFormData.classesDone)) +
                          Number(editFormData.transportAllowance) +
                          Number(editFormData.additionalBonus || 0) -
                          Number(editFormData.deductions)
                       ).toLocaleString('id-ID')}
                    </p>
                 </div>
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
                 <Button variant="ghost" onClick={() => setIsEditModalOpen(false)}>Cancel</Button>
                 <Button type="submit">Save Changes</Button>
              </div>
           </form>
         </CustomModal>
      )}

      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white mb-1">Payroll Management</h2><p className="text-gray-400 text-sm">Generate and track monthly tutor payroll.</p></div>
        <Button onClick={generatePayroll} icon={RefreshCw}>Generate Month Draft</Button>
      </div>
      <Card className="p-0 overflow-x-auto">
         <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row gap-4">
           <Input label="Month" type="select" options={MONTHS.map((m,i)=>({value:i+1, label:m}))} value={month} onChange={setMonth} className="mb-0 w-full sm:w-48" />
           <Input label="Year" type="number" value={year} onChange={setYear} className="mb-0 w-full sm:w-32" />
         </div>
         <table className="w-full text-left text-sm">
            <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
               <tr>
                  <th className="p-4 text-center w-12 text-gray-400">No.</th>
                  <th className="p-4 text-center">Tutor</th>
                  <th className="p-4 text-center">Base</th>
                  <th className="p-4 text-center">Classes</th>
                  <th className="p-4 text-center">Add. Income</th>
                  <th className="p-4 text-center">Deduct</th>
                  <th className="p-4 text-center">Total</th>
                  <th className="p-4 text-center">Status</th>
                  <th className="p-4 text-center">Actions</th>
               </tr>
            </thead>
            <tbody className="divide-y divide-gray-800">
               {paginatedPayroll.map((p, idx) => (
                  <tr key={p.id} className="hover:bg-[#0B0F19]">
                     <td className="p-4 text-center text-gray-500 font-medium">{idx + 1}</td>
                     <td className="p-4 font-bold text-white">{p.tutorName}</td>
                     <td className="p-4 text-center text-gray-300">Rp {Number(p.baseSalary).toLocaleString()}</td>
                     <td className="p-4 text-center text-gray-400">{p.classesDone}</td>
                     <td className="p-4 text-center text-green-400">+ Rp {(Number(p.transportAllowance) + Number(p.additionalBonus || 0)).toLocaleString()}</td>
                     <td className="p-4 text-center text-red-400">- Rp {Number(p.deductions).toLocaleString()}</td>
                     <td className="p-4 text-center font-bold text-[#00D4FF]">Rp {Number(p.totalPaid).toLocaleString()}</td>
                     <td className="p-4 text-center"><Badge status={p.status} /></td>
                     <td className="p-4 text-center">
                        <div className="flex justify-center gap-1 flex-wrap">
                           {p.status === 'Draft' && (
                              <>
                                 <button onClick={() => { setEditFormData(p); setIsEditModalOpen(true); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Draft"><Edit2 size={18}/></button>
                                 <button onClick={() => markAsPaid(p.id)} className="text-green-400 p-2.5 hover:bg-green-500/10 rounded-lg transition-colors" title="Mark as Paid"><CheckCircle2 size={18}/></button>
                              </>
                           )}
                           <button onClick={() => setSelectedPayroll(p)} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="View & Share Slip"><FileText size={18}/></button>
                           <button onClick={() => {
                              const tutor = db.tutors.find(t => t.id === p.tutorId);
                              const phone = tutor?.phone ? normalizeWhatsapp(tutor.phone) : '';
                              const monthName = MONTHS[Number(p.month) - 1];
                              const msg =
                                `📋 *PAYROLL SLIP — ${monthName} ${p.year}*\n` +
                                `English Club Gresik\n\n` +
                                `👤 *${p.tutorName}*\n` +
                                `📅 Periode: ${monthName} ${p.year}\n\n` +
                                `📊 *Rincian Gaji:*\n` +
                                `• Mengajar: ${p.classesDone} kelas × Rp ${Number(p.baseSalary).toLocaleString('id-ID')} = Rp ${(Number(p.baseSalary) * Number(p.classesDone)).toLocaleString('id-ID')}\n` +
                                (Number(p.transportAllowance) > 0 ? `• Reimbursement: + Rp ${Number(p.transportAllowance).toLocaleString('id-ID')}\n` : '') +
                                (Number(p.additionalBonus) > 0 ? `• Bonus Tambahan: + Rp ${Number(p.additionalBonus).toLocaleString('id-ID')}\n` : '') +
                                (Number(p.deductions) > 0 ? `• Potongan: - Rp ${Number(p.deductions).toLocaleString('id-ID')}\n` : '') +
                                `\n💰 *Take Home Pay: Rp ${Number(p.totalPaid).toLocaleString('id-ID')}*\n\n` +
                                `Status: ${p.status === 'Paid' ? '✅ Sudah Dibayar' : '⏳ Draft'}\n\n` +
                                `_Terima kasih atas dedikasi Anda! — ECG Academic Suite_`;
                              if (phone) {
                                window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, '_blank');
                              } else {
                                window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
                                showToast('No phone number saved for this tutor.', 'warning');
                              }
                           }} className="text-green-400 p-2.5 hover:bg-green-500/10 rounded-lg transition-colors" title="Send via WhatsApp"><MessageCircle size={18}/></button>
                           <button onClick={() => softDelete('payroll', p.id, `Payroll for ${p.tutorName}`)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Payroll"><Trash2 size={18}/></button>
                        </div>
                     </td>
                  </tr>
               ))}
               {paginatedPayroll.length === 0 && <tr><td colSpan={9}><EmptyState icon={DollarSign} title="No payroll generated" description="No payroll records for this month yet." /></td></tr>}
            </tbody>
         </table>
         {/* Pagination Footer */}
         <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
             <div className="flex items-center gap-2">
               <span>Show</span>
               <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                 <option value={10}>10</option>
                 <option value={20}>20</option>
                 <option value={50}>50</option>
                 <option value="All">All</option>
               </select>
               <span>entries {filteredPayroll.length > 0 && `(Total: ${filteredPayroll.length})`}</span>
             </div>

             {!isAll && totalPages > 1 && (
               <div className="flex items-center gap-2">
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                 <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
               </div>
             )}
         </div>
      </Card>
    </div>
  );
}

function CalendarModule({ db, setDb, generateId, user, showToast, softDelete }) {
  const [formData, setFormData] = useState({ id: '', date: getTodayDateLocal(), startTime: '15:00', endTime: '16:30', sessionGroup: SESSIONS[0], tutor: '', type: 'Regular Class', notes: '' });
  const [isAdding, setIsAdding] = useState(false);

  // NEW: State for Filters & Pagination
  const [filterMonth, setFilterMonth] = useState<number | string>(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  // Untuk role student: session dikunci otomatis dari data siswa (tidak bisa diubah)
  const studentSessionLocked = useMemo(() => {
    if (user.role !== 'student') return '';
    const studentRec = db.students.find(s => s.id === user.studentId);
    return studentRec ? getStudentSession(studentRec) : '';
  }, [user, db.students]);
  const [filterSession, setFilterSession] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(user.role === 'student' ? 'All' : 10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [filterMonth, filterYear, filterSession, rowsPerPage]);

  const activeTutors = db.tutors.filter(t => t.status === 'Active');

  const handleTutorToggle = (tutorName) => {
     const currentTutors = formData.tutor ? formData.tutor.split(' & ') : [];
     if (currentTutors.includes(tutorName)) {
        setFormData({...formData, tutor: currentTutors.filter(t => t !== tutorName).join(' & ')});
     } else {
        setFormData({...formData, tutor: [...currentTutors, tutorName].join(' & ')});
     }
  };

  const handleSave = (e) => {
     e.preventDefault();
     if (formData.sessionGroup !== 'All Sessions' && !formData.tutor) {
        showToast('Please select at least one tutor', 'warning');
        return;
     }
     const rec = { ...formData, id: formData.id || generateId('CAL', 'calendar'), updatedAt: getSyncTimestamp() }; // FIX BUG 3: timestamp untuk LWW
     setDb(p => ({ ...p, calendar: formData.id ? p.calendar.map(c => c.id === formData.id ? rec : c) : [...p.calendar, rec] }));
     showToast(formData.id ? 'Event updated' : 'Event created');
     setIsAdding(false);
  };

  const filteredEvents = [...db.calendar]
    .filter(c => {
       if (user.role === 'tutor' && c.sessionGroup !== 'All Sessions' && !(c.tutor && c.tutor.split(' & ').includes(user.name))) return false;
       if (user.role === 'student') {
          const studentRec = db.students.find(s => s.id === user.studentId);
          const mySession = studentRec ? getStudentSession(studentRec) : '';
          if (c.sessionGroup !== 'All Sessions' && c.sessionGroup !== mySession && c.name !== mySession) return false;
       }
       
       if (filterMonth !== 'All') {
          const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
          if (!c.date.startsWith(prefix)) return false;
       } else {
          if (!c.date.startsWith(String(filterYear))) return false;
       }
       
       const effectiveSessionFilter = user.role === 'student' ? studentSessionLocked : filterSession;
       if (effectiveSessionFilter && c.sessionGroup !== 'All Sessions' && c.sessionGroup !== effectiveSessionFilter && c.name !== effectiveSessionFilter) return false;
       
       return true;
    })
    .sort((a,b) => new Date(a.date).getTime() - new Date(b.date).getTime());

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredEvents.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredEvents.length / (rowsNum || 1));
  const paginatedData = isAll ? filteredEvents : filteredEvents.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  return (
     <div className="space-y-6">
        <div className="flex justify-between items-center">
           <div><h2 className="text-2xl font-bold text-white mb-1">Academic Calendar</h2><p className="text-gray-400 text-sm">View and manage scheduled classes and events.</p></div>
           {user.role === 'admin' && <Button onClick={() => { setFormData({ id: '', date: getTodayDateLocal(), startTime: '15:00', endTime: '16:30', sessionGroup: SESSIONS[0], tutor: '', type: 'Regular Class', notes: '' }); setIsAdding(!isAdding); }} icon={Plus}>Add Event</Button>}
        </div>
        {isAdding && user.role === 'admin' && (
           <Card className="border border-[#00D4FF]/20 shadow-[0_0_24px_rgba(0,212,255,0.06)]">
              <div className="flex items-center gap-3 mb-6 pb-4 border-b border-gray-800">
                <div className="p-2.5 rounded-xl bg-[#00D4FF]/10 border border-[#00D4FF]/20">
                  <CalendarIcon size={20} className="text-[#00D4FF]" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">{formData.id ? 'Edit Event' : 'Add New Event'}</h3>
                  <p className="text-xs text-gray-500 mt-0.5">Fields marked <span className="text-red-400">*</span> are required</p>
                </div>
              </div>
              <form onSubmit={handleSave}>
                 {/* Row 1: Date + Start + End */}
                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                   <Input label="Date" type="date" value={formData.date} onChange={v => setFormData({...formData, date: v})} required />
                   <Input label="Start Time" type="time" value={formData.startTime} onChange={v => setFormData({...formData, startTime: v})} required />
                   <Input label="End Time" type="time" value={formData.endTime} onChange={v => setFormData({...formData, endTime: v})} required />
                 </div>

                 {/* Row 2: Session + Type */}
                 <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                   <Input label="Session Group" type="select" options={['All Sessions', ...SESSIONS]} value={formData.sessionGroup} onChange={v => setFormData({...formData, sessionGroup: v})} required />
                   <Input label="Type" type="select" options={['Regular Class', 'Exam', 'Holiday', 'Meeting']} value={formData.type} onChange={v => setFormData({...formData, type: v})} required />
                 </div>

                 {/* Assign Tutors */}
                 <div className="mb-4">
                    <label className="block text-sm text-gray-400 mb-2">Assign Tutors (Multi-select) <span className="text-red-400">*</span></label>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 bg-[#0B0F19] p-3 rounded-lg border border-gray-700 max-h-40 overflow-y-auto custom-scrollbar">
                       {activeTutors.map(t => {
                          const isSelected = (formData.tutor ? formData.tutor.split(' & ') : []).includes(t.name);
                          return (
                             <label key={t.id} className={`flex items-center gap-2 p-2 rounded-lg cursor-pointer border transition-colors ${isSelected ? 'bg-[#00D4FF]/10 border-[#00D4FF]/50 text-[#00D4FF]' : 'bg-[#151B26] border-gray-800 text-gray-400 hover:border-gray-600'}`}>
                                <input type="checkbox" checked={isSelected} onChange={() => handleTutorToggle(t.name)} className="hidden" />
                                <div className={`w-4 h-4 rounded border flex items-center justify-center flex-shrink-0 ${isSelected ? 'bg-[#00D4FF] border-[#00D4FF]' : 'bg-[#0B0F19] border-gray-600'}`}>
                                   {isSelected && <Check size={12} className="text-[#0B0F19]" />}
                                </div>
                                <span className="text-sm font-medium truncate">{t.name}</span>
                             </label>
                          );
                       })}
                    </div>
                    {!formData.tutor && <p className="text-red-400 text-xs mt-1.5 flex items-center gap-1"><AlertCircle size={12}/> Please select at least one tutor.</p>}
                 </div>

                 {/* Notes */}
                 <div className="mb-6">
                   <Input label="Notes (Optional)" value={formData.notes} onChange={v => setFormData({...formData, notes: v})} placeholder="Additional notes for this event..." />
                 </div>

                 <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
                    <Button variant="secondary" onClick={() => setIsAdding(false)}>Cancel</Button>
                    <Button type="submit" icon={formData.id ? Edit2 : Plus}>{formData.id ? 'Update Event' : 'Save Event'}</Button>
                 </div>
              </form>
           </Card>
        )}
        <Card className="p-0 flex flex-col">
           {/* Filter Row */}
           <div className="p-4 bg-[#0A0E17] border-b border-gray-800 flex flex-col md:flex-row justify-between items-center gap-4">
               <div className="flex gap-2 w-full md:w-auto">
                   <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                      <option value="All">All Months</option>
                      {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
                   </select>
                   <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
               </div>
               {user.role === 'student' ? (
                  <div className="flex items-center gap-2 px-4 py-2 bg-[#00D4FF]/10 border border-[#00D4FF]/30 rounded-lg text-sm text-[#00D4FF] font-medium w-full md:w-auto">
                     <CalendarIcon size={14} className="shrink-0" />
                     <span className="truncate">{studentSessionLocked || 'My Session'}</span>
                  </div>
               ) : (
                  <select className="w-full md:w-48 bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-[#00D4FF]" value={filterSession} onChange={e => setFilterSession(e.target.value)}>
                     <option value="">All Sessions</option>
                     {SESSIONS.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
               )}
           </div>
           
           <div className="p-3 sm:p-4 bg-[#151B26] flex-1">
              <div className="max-h-[480px] overflow-y-auto custom-scrollbar pr-1 -mr-1">
              <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-2 sm:gap-3">
              {filteredEvents.map((c, idx) => {
                 const isDone = Array.from(new Set((db.studentAttendance||[]).map((a:any) => a.scheduleId))).includes(c.id);
                 const today = getTodayDateLocal();
                 const isPast = c.date < today;
                 const isToday = c.date === today;
                 const typeColors = {
                   'Holiday':  { bg: 'bg-red-500/10',    border: 'border-red-500/30',    accent: 'border-l-red-500',    text: 'text-red-400',    badge: 'bg-red-500/20 text-red-400' },
                   'Exam':     { bg: 'bg-purple-500/10', border: 'border-purple-500/30', accent: 'border-l-purple-500', text: 'text-purple-400', badge: 'bg-purple-500/20 text-purple-400' },
                   'Meeting':  { bg: 'bg-amber-500/10',  border: 'border-amber-500/30',  accent: 'border-l-amber-500',  text: 'text-amber-400',  badge: 'bg-amber-500/20 text-amber-400' },
                 };
                 const tc = typeColors[c.type] || { bg: isDone ? 'bg-emerald-950/30' : isToday ? 'bg-[#00D4FF]/5' : 'bg-[#0B0F19]', border: isDone ? 'border-emerald-700/40' : isToday ? 'border-[#00D4FF]/30' : 'border-gray-800', accent: isDone ? 'border-l-emerald-500' : isToday ? 'border-l-[#00D4FF]' : 'border-l-gray-700', text: isDone ? 'text-emerald-400' : isToday ? 'text-[#00D4FF]' : 'text-white', badge: isDone ? 'bg-emerald-500/20 text-emerald-400' : 'bg-blue-500/20 text-blue-400' };
                 const numColor = isDone ? 'bg-emerald-500/20 border-emerald-500/30 text-emerald-400' : isToday ? 'bg-[#00D4FF]/15 border-[#00D4FF]/30 text-[#00D4FF]' : 'bg-white/5 border-white/10 text-gray-400';
                 return (
                   <div key={c.id} className={`relative rounded-xl border border-l-4 ${tc.bg} ${tc.border} ${tc.accent} p-3 flex flex-col gap-2 transition-all hover:brightness-110 ${isPast && !isDone && c.type === 'Regular Class' ? 'opacity-60' : ''}`}
                     style={{backdropFilter:'blur(8px)', WebkitBackdropFilter:'blur(8px)'}}>
                     {/* Top row: number + session name + badges */}
                     <div className="flex items-start gap-2">
                       <div className={`flex-shrink-0 w-7 h-7 rounded-full border flex items-center justify-center ${numColor}`}>
                         <span className="text-[10px] font-black">{String(idx + 1).padStart(2, '0')}</span>
                       </div>
                       <div className="flex-1 min-w-0">
                         <p className={`text-sm font-bold leading-tight truncate ${tc.text}`}>{c.sessionGroup || c.name}</p>
                         <div className="flex flex-wrap gap-1 mt-1">
                           <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wide ${tc.badge}`}>{c.type}</span>
                           {isToday && !isDone && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded uppercase bg-[#00D4FF]/20 text-[#00D4FF] tracking-wide">Today</span>}
                           {isDone && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded uppercase bg-emerald-500/20 text-emerald-400 tracking-wide flex items-center gap-0.5"><CheckCircle2 size={9}/>Done</span>}
                         </div>
                       </div>
                       {user.role === 'admin' && (
                         <div className="flex gap-1 flex-shrink-0">
                           <button onClick={() => { setFormData(c); setIsAdding(true); }} className="text-gray-500 hover:text-blue-400 p-1.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit"><Edit2 size={14}/></button>
                           <button onClick={() => softDelete('calendar', c.id, 'Calendar Event')} className="text-gray-500 hover:text-red-400 p-1.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete"><Trash2 size={14}/></button>
                         </div>
                       )}
                     </div>
                     {/* Bottom row: date, time, tutor */}
                     <div className="flex flex-wrap gap-x-3 gap-y-0.5 text-[11px] text-gray-400 pl-9">
                       <span className="flex items-center gap-1"><CalendarIcon size={10}/>{safeDateDisplay(c.date, 'en-GB', {weekday:'short', day:'numeric', month:'short', year:'numeric'})}</span>
                       {c.startTime && <span className="flex items-center gap-1"><Clock size={10}/>{c.startTime}{c.endTime ? ` - ${c.endTime}` : ''}</span>}
                       {c.tutor && <span className="flex items-center gap-1 truncate max-w-full"><User size={10}/><span className="truncate">{c.tutor}</span></span>}
                       {c.notes && <span className="italic text-gray-600 truncate max-w-full">{c.notes}</span>}
                     </div>
                   </div>
                 );
              })}
              {filteredEvents.length === 0 && <div className="col-span-full text-center text-gray-500 py-8">No events scheduled matching your filters.</div>}
              </div>
              </div>
           </div>
        </Card>
     </div>
  );
}

function AnnouncementsModule({ db, setDb, generateId, user, showToast, softDelete, setActiveTab }) {
  const [formData, setFormData] = useState({ id: '', title: '', content: '' });
  const [isAdding, setIsAdding] = useState(false);

  // NEW: State for Pagination
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [rowsPerPage]);

  const handleSave = (e) => {
     e.preventDefault();
     const rec = { ...formData, id: formData.id || generateId('ANN', 'announcements'), date: getTodayDateLocal(), author: user.name, updatedAt: getSyncTimestamp() }; // FIX BUG 3: timestamp untuk LWW
     setDb(p => ({ ...p, announcements: formData.id ? p.announcements.map(a => a.id === formData.id ? rec : a) : [...p.announcements, rec] }));
     showToast('Announcement published');
     setIsAdding(false);
     setFormData({ id: '', title: '', content: '' });
  };

  const sortedAnnouncements = [...db.announcements].reverse();

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? sortedAnnouncements.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(sortedAnnouncements.length / (rowsNum || 1));
  const paginatedData = isAll ? sortedAnnouncements : sortedAnnouncements.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  return (
     <div className="space-y-6">
        <div className="flex justify-between items-center">
           <div><h2 className="text-2xl font-bold text-white mb-1">Announcements</h2><p className="text-gray-400 text-sm">Post and manage announcements visible to students and tutors.</p></div>
           {user.role === 'admin' && <Button onClick={() => { setIsAdding(!isAdding); setFormData({ id: '', title: '', content: '' }); }} icon={Plus}>New Post</Button>}
        </div>
        {isAdding && user.role === 'admin' && (
           <Card>
              <form onSubmit={handleSave} className="space-y-4">
                 <Input label="Title" value={formData.title} onChange={v => setFormData({...formData, title: v})} required />
                 <div>
                    <label className="block text-sm text-gray-400 mb-1">Content</label>
                    <textarea className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg p-3 text-white focus:border-[#00D4FF]" rows={4} value={formData.content} onChange={e => setFormData({...formData, content: e.target.value})} required></textarea>
                 </div>
                 <div className="flex justify-end gap-2 mt-4">
                    <Button variant="ghost" onClick={() => setIsAdding(false)}>Cancel</Button>
                    <Button type="submit">Publish</Button>
                 </div>
              </form>
           </Card>
        )}
        <div className="space-y-4">
           {paginatedData.map((a, idx) => (
              <GlassCard key={a.id} accentColor="yellow" className="border-l-4 border-l-yellow-400">
                 <div className="flex justify-between items-start mb-3">
                    <div className="flex items-start gap-3">
                       <span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-yellow-400/10 border border-yellow-400/30 text-[11px] font-black text-yellow-400 shrink-0 mt-0.5">{String(idx + 1).padStart(2, '0')}</span>
                       <div>
                          <h3 className="text-xl font-bold text-white">{a.title}</h3>
                          <p className="text-xs text-gray-500 uppercase tracking-widest mt-1">{a.date} • By {a.author}</p>
                       </div>
                    </div>
                    {user.role === 'admin' && (
                       <div className="flex gap-2 shrink-0">
                          <button onClick={() => { setFormData(a); setIsAdding(true); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Post"><Edit2 size={18}/></button>
                          <button onClick={() => softDelete('announcements', a.id, 'Announcement')} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Post"><Trash2 size={18}/></button>
                       </div>
                    )}
                 </div>
                 <p className="text-gray-300 whitespace-pre-wrap leading-relaxed text-sm">{a.content}</p>
              </GlassCard>
           ))}
           {paginatedData.length === 0 && <p className="text-center text-gray-500 bg-[#151B26] p-8 rounded-xl border border-gray-800">No announcements yet.</p>}
        </div>

        {/* Pagination Footer */}
        {sortedAnnouncements.length > 0 && (
          <div className="p-4 bg-[#0A0E17] border border-gray-800 rounded-xl flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400 mt-4">
             <div className="flex items-center gap-2">
               <span>Show</span>
               <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                 <option value={10}>10</option>
                 <option value={20}>20</option>
                 <option value={50}>50</option>
                 <option value="All">All</option>
               </select>
               <span>entries (Total: {sortedAnnouncements.length})</span>
             </div>
             
             {!isAll && totalPages > 1 && (
               <div className="flex items-center gap-2">
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                 <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
               </div>
             )}
          </div>
        )}
     </div>
  );
}

// MODULE UNTUK BUAT AKUN TERMASUK SISWA
function SettingsModule({ db, setDb, generateId, user, showToast, requestConfirm, getAuthToken, dbVersion, currentUser, setSyncStatus }) {
  const [formData, setFormData] = useState({ name: '', username: '', password: '', role: 'admin', active: 'Active', studentId: '', tutorId: '', teachingSession: '' });
  const [isEditingId, setIsEditingId] = useState(null);
  const [resetDialog, setResetDialog] = useState(null);
  const [newPassword, setNewPassword] = useState('');
  const [credentialCard, setCredentialCard] = useState(null); // { name, username, password, role }

  // NEW: State for Search, Filter & Pagination
  const [searchTerm, setSearchTerm] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [filterRole, setFilterRole] = useState('All');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    const handler = setTimeout(() => setDebouncedSearch(searchTerm), 300);
    return () => clearTimeout(handler);
  }, [searchTerm]);

  useEffect(() => {
    setCurrentPage(1);
  }, [debouncedSearch, filterRole, rowsPerPage]);

  // PERBAIKAN: Membungkus filter dengan sortStudentsLogically
  const activeStudentsWithoutAccount = sortStudentsLogically(db.students.filter(s => s.status === 'Active' && !db.users.some(u => u.role === 'student' && u.studentId === s.id)));
  const activeTutorsWithoutAccount = db.tutors.filter(t => t.status === 'Active' && (!t.username || !t.password));

  // NEW: Gabungkan admin/tutor/student ke satu array terpadu untuk search, filter, & paginasi
  const unifiedUsers = useMemo(() => {
    const admins = db.users.filter(u => u.role !== 'student').map(u => {
      const isSuperAdmin = (u.username || '').toLowerCase() === 'vicky' || String(u.role).toLowerCase().includes('super');
      return { key: `admin-${u.id}`, roleLabel: isSuperAdmin ? 'Super Admin' : 'Admin', roleType: isSuperAdmin ? 'super' : 'admin', name: u.name, username: u.username, hasAccount: true, status: u.active, raw: u };
    });
    const tutors = db.tutors.map(t => {
      const hasAccount = !!(t.username && t.password);
      return { key: `tutor-${t.id}`, roleLabel: 'Tutor', roleType: 'tutor', name: t.name, username: hasAccount ? t.username : '', hasAccount, status: t.status, raw: t };
    });
    const students = db.students.map(s => {
      const userAcc = db.users.find(u => u.role === 'student' && (u.studentId === s.id || u.name === s.name));
      return { key: `student-${s.id}`, roleLabel: 'Student', roleType: 'student', name: s.name, username: userAcc ? userAcc.username : '', hasAccount: !!userAcc, status: userAcc ? userAcc.active : null, raw: s, userAcc };
    });
    const orphanStudents = db.users.filter(u => u.role === 'student' && !db.students.some(s => s.id === u.studentId || s.name === u.name)).map(u => ({
      key: `orphan-${u.id}`, roleLabel: 'Student', roleType: 'student', name: u.name, username: u.username, hasAccount: true, status: u.active, raw: u, isOrphan: true
    }));
    return [...admins, ...tutors, ...students, ...orphanStudents];
  }, [db.users, db.tutors, db.students]);

  const filteredUsers = useMemo(() => {
    return unifiedUsers.filter(u => {
      if (filterRole !== 'All' && u.roleType !== filterRole.toLowerCase()) return false;
      if (debouncedSearch) {
        const term = debouncedSearch.toLowerCase();
        if (!u.name.toLowerCase().includes(term) && !(u.username || '').toLowerCase().includes(term)) return false;
      }
      return true;
    });
  }, [unifiedUsers, filterRole, debouncedSearch]);

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredUsers.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredUsers.length / (rowsNum || 1));
  const paginatedUsers = isAll ? filteredUsers : filteredUsers.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  const handleSave = (e) => {
    e.preventDefault();
    if (formData.password && formData.password.length < 6) {
      return showToast('Password minimal 6 karakter. Kosongkan field untuk auto-generate password acak.', 'warning');
    }
    if (formData.role === 'tutor') {
      if (!isEditingId && !formData.tutorId) return showToast('Please select a tutor to link', 'error');
      const targetId = isEditingId || formData.tutorId;
      const linkedTutor = db.tutors.find(t => t.id === targetId);
      const _genTutorPwd = () => { const c='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#'; let p=''; for(let i=0;i<12;i++) p+=c[Math.floor(Math.random()*c.length)]; return p; };
      const _tutorTempPwd = _genTutorPwd();
      const savedTutorPassword = formData.password !== '' ? String(formData.password).trim() : (linkedTutor?.password ? String(linkedTutor.password).trim() : _tutorTempPwd);
      if (!isEditingId && !formData.password) {
        showToast(`Password tidak diisi. Akun Tutor dibuat dengan password sementara: ${_tutorTempPwd} — Catat & berikan ke tutor!`, 'warning');
      }
      setDb((p) => ({ 
        ...p, 
        // BUG FIX E: Pastikan password lama tidak hilang (undefined) saat edit.
        // Jika admin kosongkan field password saat edit → pakai password lama.
        // Jika tutor belum punya password sama sekali → pakai password baru yang diisi admin.
        tutors: p.tutors.map((t) => (t.id === targetId ? { 
          ...t, 
          username: String(formData.username || '').trim(),
          password: savedTutorPassword,
          status: formData.active || t.status || 'Active',
          mustChangePassword: !isEditingId ? true : t.mustChangePassword,
          updatedAt: getSyncTimestamp() // BUGFIX: Tambahkan updatedAt agar LWW tidak kalah ke cloud lama
        } : t)) 
      }));
      if (!isEditingId) {
        setCredentialCard({ name: linkedTutor?.name || formData.username, username: String(formData.username).trim(), password: savedTutorPassword, role: 'Tutor' });
      }
    } else if (formData.role === 'student') {
      if (!formData.studentId) return showToast('Please select a student to link', 'error');
      const linkedStudent = db.students.find(s => s.id === formData.studentId) || { name: formData.name };
      const existingUser = db.users.find(u => u.id === isEditingId);
      // Pastikan password tidak pernah kosong:
      // - Edit mode: boleh kosong di form (artinya pakai password lama)
      // - Create mode: wajib isi; kalau tetap kosong, generate random temp password
      const generateTempPassword = () => {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#';
        let pwd = '';
        for (let i = 0; i < 12; i++) pwd += chars[Math.floor(Math.random() * chars.length)];
        return pwd;
      };
      const tempPwd = generateTempPassword();
      const finalPassword = formData.password !== ''
        ? formData.password
        : (existingUser ? existingUser.password : tempPwd);
      if (!isEditingId && !formData.password) {
        showToast(`Password tidak diisi. Akun dibuat dengan password sementara: ${finalPassword} — Catat & berikan ke siswa!`, 'warning');
      }
      // BUGFIX: Tambahkan updatedAt agar LWW tidak kalah ke cloud lama
      const rec = { ...formData, name: linkedStudent.name, password: finalPassword, id: isEditingId || generateId('USR', 'users'), mustChangePassword: !isEditingId, updatedAt: getSyncTimestamp() };
      setDb((p) => ({ ...p, users: isEditingId ? p.users.map((u) => (u.id === isEditingId ? rec : u)) : [...p.users, rec] }));
      if (!isEditingId) {
        setCredentialCard({ name: linkedStudent.name, username: String(formData.username).trim(), password: finalPassword, role: 'Student' });
      }
    } else {
      // ADMIN / SUPER ADMIN (Fallback aman untuk setiap role selain student dan tutor)
      const existingUser = db.users.find(u => u.id === isEditingId);
      const finalPassword = formData.password !== '' ? formData.password : (existingUser ? existingUser.password : '');
      // BUGFIX: Tambahkan updatedAt agar LWW di mergeByIds selalu memenangkan versi lokal
      // yang baru diedit. Tanpa updatedAt, parseTime() = 0 → cloud selalu menang → nama/data
      // yang baru diubah (termasuk nama super admin) akan kembali ke versi lama dari cloud.
      const rec = { ...formData, password: finalPassword, id: isEditingId || generateId('ADM', 'users'), mustChangePassword: !isEditingId, updatedAt: getSyncTimestamp() };
      setDb((p) => ({ ...p, users: isEditingId ? p.users.map((u) => (u.id === isEditingId ? rec : u)) : [...p.users, rec] }));
    }
    showToast('User Saved');
    setIsEditingId(null);
    setFormData({ name: '', username: '', password: '', role: 'admin', active: 'Active', studentId: '', tutorId: '', teachingSession: '' });
  };

  const handleResetPassword = (e) => {
    e.preventDefault();
    if(!newPassword) return showToast('Please enter a new password', 'warning');
    if(newPassword.length < 6) return showToast('Password minimal 6 karakter.', 'warning');

    const targetUsername = resetDialog.username || '';
    
    // Update db state lokal
    let updatedDb;
    if (resetDialog.role !== 'tutor') {
      setDb(p => {
        updatedDb = {...p, users: p.users.map(u => u.id === resetDialog.id ? {...u, password: newPassword, mustChangePassword: true} : u)};
        return updatedDb;
      });
    } else {
      setDb(p => {
        updatedDb = {...p, tutors: p.tutors.map(t => t.id === resetDialog.id ? {...t, password: newPassword, mustChangePassword: true} : t)};
        return updatedDb;
      });
    }

    // Tidak ada force-push manual di sini. Auto-sync menggunakan latestDbRef + delta
    // sehingga reset sandi tidak pernah mengirim snapshot db lama dari closure.

    showToast(`Password reset for ${resetDialog.name} was successful.`);
    setCredentialCard({ name: resetDialog.name, username: resetDialog.username || '', password: newPassword, role: resetDialog.role === 'tutor' ? 'Tutor' : resetDialog.role === 'student' ? 'Student' : 'Admin' });
    setResetDialog(null);
    setNewPassword('');
  };

  const handleDeleteUser = (id, role, name) => {
    const targetUser = db.users.find(u => u.id === id);
    
    // Perlindungan: cegah SEMUA super admin terhapus habis (mencegah lockout total dari sistem).
    // Selama masih ada minimal 1 akun super admin lain, akun super admin duplikat/lama BOLEH dihapus.
    if (targetUser && String(targetUser.role).toLowerCase().includes('super')) {
      const remainingSuperAdmins = db.users.filter(u => String(u.role).toLowerCase().includes('super') && u.id !== id).length;
      if (remainingSuperAdmins < 1) {
        showToast('Cannot delete the last remaining Super Admin. Create/keep at least one other Super Admin first.', 'error');
        return;
      }
    }
    
    requestConfirm(
      'Confirm Account Deletion',
      role === 'tutor'
         ? `Are you sure you want to remove the login access for ${name}? Their profile will remain in the Tutors Directory.`
         : `Are you sure you want to delete the account for ${name}? This action cannot be undone.`,
      () => {
        if (role !== 'tutor') {
          setDb(p => ({ ...p, users: p.users.filter(u => u.id !== id) }));
        } else {
          // PERBAIKAN: Hanya hapus kredensial login, JANGAN hapus profil dari db.tutors
          setDb(p => ({ ...p, tutors: p.tutors.map(t => t.id === id ? { ...t, username: '', password: '' } : t) }));
        }
        showToast('Account removed successfully');
      }
    );
  };

  // Helper: generate random temp password (reused also in quick-create buttons below)
  const generateTempPassword = () => {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#';
    let pwd = '';
    for (let i = 0; i < 12; i++) pwd += chars[Math.floor(Math.random() * chars.length)];
    return pwd;
  };

  return (
    <div className="space-y-6">

      {/* ============================================================
          CREDENTIAL CARD MODAL
          Muncul otomatis setelah admin buat/reset akun siswa atau tutor.
          Admin bisa salin teks atau print langsung untuk diserahkan ke siswa/tutor.
          ============================================================ */}
      <CustomModal isOpen={!!credentialCard} onClose={() => setCredentialCard(null)} title="🎉 Account Created Successfully">
        {credentialCard && (
          <div className="space-y-5">
            <p className="text-gray-400 text-sm">Here are the login credentials for <span className="text-white font-semibold">{credentialCard.name}</span>. Copy or print them and hand them directly to the person.</p>

            {/* Printable credential card */}
            <div id="credential-print-area" className="bg-[#0B0F19] border border-[#00D4FF]/30 rounded-xl p-5 space-y-4">
              <div className="flex items-center gap-3 border-b border-gray-800 pb-3">
                <div className="p-2 rounded-lg bg-[#00D4FF]/10 border border-[#00D4FF]/20">
                  <ShieldCheck size={20} className="text-[#00D4FF]" />
                </div>
                <div>
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Login Credentials</p>
                  <p className="text-white font-bold">{credentialCard.name} <span className="text-xs font-normal text-[#00D4FF] ml-1">({credentialCard.role})</span></p>
                </div>
              </div>

              <div className="space-y-3">
                <div className="flex items-center justify-between bg-[#151B26] rounded-lg px-4 py-3">
                  <div>
                    <p className="text-[11px] text-gray-500 uppercase tracking-wider mb-0.5">Username</p>
                    <p className="text-white font-mono font-bold text-lg">{credentialCard.username}</p>
                  </div>
                  <button
                    onClick={() => { navigator.clipboard?.writeText(credentialCard.username); showToast('Username copied!', 'success'); }}
                    className="p-2 hover:bg-gray-700 rounded-lg transition-colors text-gray-400 hover:text-white"
                    title="Copy username"
                  ><Copy size={16} /></button>
                </div>

                <div className="flex items-center justify-between bg-[#151B26] rounded-lg px-4 py-3">
                  <div>
                    <p className="text-[11px] text-gray-500 uppercase tracking-wider mb-0.5">Temporary Password</p>
                    <p className="text-white font-mono font-bold text-lg tracking-widest">{credentialCard.password}</p>
                  </div>
                  <button
                    onClick={() => { navigator.clipboard?.writeText(credentialCard.password); showToast('Password copied!', 'success'); }}
                    className="p-2 hover:bg-gray-700 rounded-lg transition-colors text-gray-400 hover:text-white"
                    title="Copy password"
                  ><Copy size={16} /></button>
                </div>
              </div>

              <div className="bg-amber-500/10 border border-amber-500/20 rounded-lg px-4 py-3 flex gap-2.5 items-start">
                <AlertCircle size={16} className="text-amber-400 mt-0.5 shrink-0" />
                <p className="text-amber-300 text-xs leading-relaxed">
                  {credentialCard.role === 'Student' || credentialCard.role === 'Tutor'
                    ? `${credentialCard.name} will be prompted to change this password on first login.`
                    : 'Make sure these credentials are delivered in person or through a secure channel.'}
                </p>
              </div>
            </div>

            {/* Action buttons */}
            <div className="flex flex-col sm:flex-row gap-3 pt-1">
              <button
                onClick={() => {
                  const text = `LOGIN CREDENTIALS\n${'─'.repeat(30)}\nName    : ${credentialCard.name}\nRole    : ${credentialCard.role}\nUsername: ${credentialCard.username}\nPassword: ${credentialCard.password}\n${'─'.repeat(30)}\n* Please change your password after first login.`;
                  navigator.clipboard?.writeText(text);
                  showToast('Full credentials copied to clipboard!', 'success');
                }}
                className="flex-1 flex items-center justify-center gap-2 bg-[#00D4FF]/10 hover:bg-[#00D4FF]/20 border border-[#00D4FF]/30 text-[#00D4FF] rounded-xl px-4 py-3 text-sm font-semibold transition-colors"
              >
                <Copy size={16} /> Copy All
              </button>
              <button
                onClick={() => {
                  const w = window.open('', '_blank');
                  w.document.write(`<!DOCTYPE html><html><head><title>Login Credentials - ${credentialCard.name}</title><style>
                    body{font-family:Arial,sans-serif;background:#fff;color:#111;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
                    .card{border:2px solid #00D4FF;border-radius:16px;padding:32px 40px;max-width:400px;width:100%;box-shadow:0 4px 24px rgba(0,212,255,0.15)}
                    h2{color:#00D4FF;margin:0 0 4px;font-size:18px}
                    .sub{color:#666;font-size:13px;margin:0 0 24px}
                    .field{background:#f5f5f5;border-radius:10px;padding:12px 16px;margin-bottom:12px}
                    .label{font-size:11px;color:#999;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px}
                    .value{font-family:monospace;font-size:20px;font-weight:700;color:#111;letter-spacing:.08em}
                    .note{background:#fffbe6;border:1px solid #f59e0b;border-radius:8px;padding:10px 14px;font-size:12px;color:#92400e;margin-top:16px}
                    .divider{border:none;border-top:1px solid #eee;margin:20px 0}
                    .logo{font-size:13px;color:#999;text-align:center;margin-top:20px}
                  </style></head><body>
                  <div class="card">
                    <h2>🔐 Login Credentials</h2>
                    <p class="sub">${credentialCard.name} &bull; ${credentialCard.role}</p>
                    <div class="field"><div class="label">Username</div><div class="value">${credentialCard.username}</div></div>
                    <div class="field"><div class="label">Temporary Password</div><div class="value">${credentialCard.password}</div></div>
                    <div class="note">⚠️ Please change your password after your first login for account security.</div>
                    <hr class="divider"/>
                    <div class="logo">ECG Academic Suite</div>
                  </div>
                  <script>window.onload=()=>window.print()</script>
                  </body></html>`);
                  w.document.close();
                }}
                className="flex-1 flex items-center justify-center gap-2 bg-purple-500/10 hover:bg-purple-500/20 border border-purple-500/30 text-purple-300 rounded-xl px-4 py-3 text-sm font-semibold transition-colors"
              >
                <Printer size={16} /> Print / Save PDF
              </button>
              <button
                onClick={() => setCredentialCard(null)}
                className="sm:w-auto flex items-center justify-center gap-2 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 rounded-xl px-4 py-3 text-sm font-semibold transition-colors"
              >
                <Check size={16} /> Done
              </button>
            </div>
          </div>
        )}
      </CustomModal>

      <CustomModal isOpen={!!resetDialog} onClose={() => {setResetDialog(null); setNewPassword('');}} title="Reset Password">
        <p className="text-gray-400 mb-6 text-sm">You are manually resetting the password for <b>{resetDialog?.name}</b>.</p>
        <form onSubmit={handleResetPassword} className="space-y-4">
           <Input label="New Password" type="password" value={newPassword} onChange={setNewPassword} required placeholder="Enter new password" />
           <div className="flex justify-end gap-3 pt-2">
             <Button variant="ghost" onClick={() => {setResetDialog(null); setNewPassword('');}}>Cancel</Button>
             <Button type="submit" className="bg-red-500 hover:bg-red-600 text-white border-none shadow-none">Force Reset</Button>
           </div>
        </form>
      </CustomModal>

      <Card className="border border-amber-500/20 shadow-[0_0_24px_rgba(245,158,11,0.06)]">
        <div className="flex items-center gap-3 mb-6 pb-4 border-b border-gray-800">
          <div className="p-2.5 rounded-xl bg-amber-500/10 border border-amber-500/20">
            <Shield size={20} className="text-amber-400" />
          </div>
          <div>
            <h3 className="text-base font-bold text-white">{isEditingId ? 'Edit User Account' : 'Create User Account'}</h3>
            <p className="text-xs text-gray-500 mt-0.5">Fields marked <span className="text-red-400">*</span> are required</p>
          </div>
        </div>
        <form onSubmit={handleSave}>
          {/* Row 1: Role + Name/Link */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <Input label="Role" type="select" options={formData.role === 'super admin' ? ['super admin', 'admin', 'tutor', 'student'] : ['admin', 'tutor', 'student']} value={formData.role} onChange={(v) => setFormData({ ...formData, role: v })} required disabled={isEditingId !== null} />
            {formData.role === 'student' && !isEditingId ? (
               <Input label="Link to Student" type="select" options={activeStudentsWithoutAccount.map(s => ({ value: s.id, label: `${s.name} (${s.class})` }))} value={formData.studentId} onChange={(v) => setFormData({ ...formData, studentId: v })} required />
            ) : formData.role === 'tutor' && !isEditingId ? (
               <Input label="Link to Tutor" type="select" options={activeTutorsWithoutAccount.map(t => ({ value: t.id, label: t.name }))} value={formData.tutorId} onChange={(v) => setFormData({ ...formData, tutorId: v })} required />
            ) : (
               <Input label="Full Name" value={formData.name} onChange={(v) => setFormData({ ...formData, name: v })} required disabled={formData.role === 'student' || formData.role === 'tutor'} />
            )}
          </div>
          {/* Row 2: Username + Password + Status */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <Input label="Username" value={formData.username} onChange={(v) => setFormData({ ...formData, username: v })} required />
            <Input label="Password" type="text" value={formData.password} onChange={(v) => setFormData({ ...formData, password: v })} placeholder={isEditingId ? 'Leave blank to keep current password' : 'Leave blank to auto-generate'} />
            <Input label="Status" type="select" options={['Active', 'Inactive']} value={formData.active} onChange={(v) => setFormData({ ...formData, active: v })} required />
          </div>
          <div className="flex justify-end gap-3 pt-4 border-t border-gray-800">
            {isEditingId && <Button variant="secondary" onClick={() => { setIsEditingId(null); setFormData({ name: '', username: '', password: '', role: 'admin', active: 'Active', studentId: '', tutorId: '', teachingSession: '' }); }}>Cancel</Button>}
            <Button type="submit" icon={isEditingId ? Edit2 : Plus}>{isEditingId ? 'Update Account' : 'Save Account'}</Button>
          </div>
        </form>
      </Card>
      
      <Card className="p-0 flex flex-col">
        <div className="p-4 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row gap-3 justify-between items-center">
           <div className="relative w-full sm:w-64">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500" />
              <input
                type="text"
                placeholder="Search name or username..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
                className="w-full bg-[#151B26] border border-gray-700 rounded-lg pl-9 pr-3 py-2 text-sm text-white focus:outline-none focus:border-[#00D4FF]"
              />
           </div>
           <select
              className="w-full sm:w-48 bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-[#00D4FF]"
              value={filterRole}
              onChange={e => setFilterRole(e.target.value)}
           >
              <option value="All">All Roles</option>
              <option value="admin">Admin</option>
              <option value="tutor">Tutor</option>
              <option value="student">Student</option>
           </select>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">Role</th><th className="p-4 text-center">Name</th><th className="p-4 text-center">Username</th><th className="p-4 text-center">Password</th><th className="p-4 text-center">Status</th><th className="p-4 text-center">Actions</th></tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {paginatedUsers.map((row, idx) => {
              const isSuperAdmin = row.roleType === 'super';
              const globalIdx = (currentPage - 1) * (isAll ? filteredUsers.length : rowsNum) + idx + 1;
              return (
                <tr key={row.key} className="hover:bg-[#0B0F19]">
                  <td className="p-4 text-center text-gray-500 font-medium">{globalIdx}</td>
                  <td className={`p-4 text-center text-xs font-bold uppercase ${isSuperAdmin ? 'text-red-500' : row.roleType === 'tutor' ? 'text-purple-400' : 'text-blue-400'}`}>
                    {row.roleLabel}{row.isOrphan && <span className="text-red-500"> *</span>}
                  </td>
                  <td className="p-4 text-center text-white">{row.name}{row.isOrphan && <span className="text-[11px] block text-red-400">(Unlinked)</span>}</td>
                  <td className="p-4 text-center">{row.hasAccount ? row.username : <span className="text-gray-600 text-xs italic">No Account</span>}</td>
                  <td className="p-4 text-center font-mono text-gray-400">{row.hasAccount ? '••••••••' : '-'}</td>
                  <td className="p-4 text-center">
                    {row.hasAccount ? <Badge status={row.status} /> : <span className="px-2.5 py-1 text-xs font-medium rounded-full border bg-gray-500/10 text-gray-400 border-gray-500/20 whitespace-nowrap">No Account</span>}
                  </td>
                  <td className="p-4 text-center flex justify-center gap-2">
                    {isSuperAdmin ? (
                      <>
                        <button onClick={() => { setFormData({ ...row.raw }); setIsEditingId(row.raw.id); const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Super Admin"><UserCog size={18} /></button>
                        <button onClick={() => { setResetDialog({ id: row.raw.id, role: row.raw.role, name: row.raw.name, username: row.raw.username }); }} className="text-yellow-500 p-2.5 hover:bg-yellow-500/10 rounded-lg transition-colors" title="Reset Password"><KeyRound size={18} /></button>
                        <button onClick={() => handleDeleteUser(row.raw.id, row.raw.role, row.raw.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Super Admin (hanya bisa jika masih ada Super Admin lain)"><Trash2 size={18} /></button>
                      </>
                    ) : row.roleType === 'admin' ? (
                      <>
                        <button onClick={() => { setFormData({ ...row.raw }); setIsEditingId(row.raw.id); const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Profile"><UserCog size={18} /></button>
                        <button onClick={() => { setResetDialog({ id: row.raw.id, role: row.raw.role, name: row.raw.name, username: row.raw.username }); }} className="text-yellow-500 p-2.5 hover:bg-yellow-500/10 rounded-lg transition-colors" title="Reset Password"><KeyRound size={18} /></button>
                        <button onClick={() => handleDeleteUser(row.raw.id, row.raw.role, row.raw.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete User"><Trash2 size={18} /></button>
                      </>
                    ) : row.roleType === 'tutor' ? (
                      <>
                        <button onClick={() => {
                          if (row.hasAccount) { setFormData({ ...row.raw, role: 'tutor', active: row.raw.status, tutorId: row.raw.id, studentId: '' }); setIsEditingId(row.raw.id); }
                          else { setFormData({ name: row.raw.name, username: row.raw.name.toLowerCase().replace(/[^a-z0-9]/g, ''), password: generateTempPassword(), role: 'tutor', active: row.raw.status || 'Active', tutorId: row.raw.id, studentId: '', teachingSession: row.raw.teachingSession || '' }); setIsEditingId(null); }
                          const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50);
                        }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title={row.hasAccount ? "Edit Account" : "Create Account"}><UserCog size={18} /></button>
                        {row.hasAccount && (
                           <>
                              <button onClick={() => setResetDialog({ id: row.raw.id, role: 'tutor', name: row.raw.name, username: row.raw.username })} className="text-yellow-500 p-2.5 hover:bg-yellow-500/10 rounded-lg transition-colors" title="Reset Password"><KeyRound size={18} /></button>
                              <button onClick={() => handleDeleteUser(row.raw.id, 'tutor', row.raw.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Remove Account"><Trash2 size={18} /></button>
                           </>
                        )}
                      </>
                    ) : row.isOrphan ? (
                      <>
                        <button onClick={() => { setFormData({ ...row.raw }); setIsEditingId(row.raw.id); const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50); }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Profile"><UserCog size={18} /></button>
                        <button onClick={() => { setResetDialog({ id: row.raw.id, role: row.raw.role, name: row.raw.name, username: row.raw.username }); }} className="text-yellow-500 p-2.5 hover:bg-yellow-500/10 rounded-lg transition-colors" title="Reset Password"><KeyRound size={18} /></button>
                        <button onClick={() => handleDeleteUser(row.raw.id, row.raw.role, row.raw.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete User"><Trash2 size={18} /></button>
                      </>
                    ) : (
                      <>
                        <button onClick={() => {
                          if (row.userAcc) { setFormData({ tutorId: '', teachingSession: '', ...row.userAcc, studentId: row.raw.id }); setIsEditingId(row.userAcc.id); }
                          else { setFormData({ name: row.raw.name, username: row.raw.name.toLowerCase().replace(/[^a-z0-9]/g, ''), password: generateTempPassword(), role: 'student', active: 'Active', studentId: row.raw.id, tutorId: '', teachingSession: '' }); setIsEditingId(null); }
                          const contentEl = document.querySelector('main'); setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50);
                        }} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title={row.userAcc ? "Edit Account" : "Create Account"}><UserCog size={18} /></button>
                        {row.userAcc && (
                           <>
                              <button onClick={() => setResetDialog({ id: row.userAcc.id, role: row.userAcc.role, name: row.userAcc.name, username: row.userAcc.username })} className="text-yellow-500 p-2.5 hover:bg-yellow-500/10 rounded-lg transition-colors" title="Reset Password"><KeyRound size={18} /></button>
                              <button onClick={() => handleDeleteUser(row.userAcc.id, row.userAcc.role, row.userAcc.name)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Account"><Trash2 size={18} /></button>
                           </>
                        )}
                      </>
                    )}
                  </td>
                </tr>
              );
            })}
            {paginatedUsers.length === 0 && <tr><td colSpan={7}><EmptyState icon={Users} title="No users found" description="Try adjusting your search or filter." /></td></tr>}
          </tbody>
        </table>
        </div>
        <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
            <div className="flex items-center gap-2">
              <span>Show</span>
              <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                <option value={10}>10</option>
                <option value={20}>20</option>
                <option value={50}>50</option>
                <option value="All">All</option>
              </select>
              <span>entries (Total: {filteredUsers.length})</span>
            </div>

            {!isAll && totalPages > 1 && (
              <div className="flex items-center gap-2">
                <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
              </div>
            )}
        </div>
      </Card>
    </div>
  );
}

function AccountSettingsModule({ db, setDb, user, setCurrentUser, showToast, language = 'en' }) {
  const [currentPwd, setCurrentPwd] = useState('');
  const [newPwd, setNewPwd] = useState('');
  const [confirmPwd, setConfirmPwd] = useState('');

  const handleUpdate = (e) => {
    e.preventDefault();

    const currentNumStr = (!isNaN(Number(currentPwd)) && currentPwd !== '') ? String(Number(currentPwd)) : null;
    // FIX #1: Jika GAS mengirim '__MASKED__' sebagai sentinel (password di-mask server),
    // lewati pengecekan password lama agar siswa/tutor tetap bisa ganti password.
    const isMaskedPassword = !user.password || user.password === '__MASKED__';
    if (!isMaskedPassword && String(currentPwd) !== String(user.password) && (!currentNumStr || String(user.password) !== currentNumStr)) {
      return showToast('Your current password is incorrect.', 'error');
    }
    if (newPwd !== confirmPwd) {
      return showToast('New password and confirmation do not match.', 'error');
    }
    if (newPwd.length < 6) {
      return showToast('Password must be at least 6 characters long.', 'warning');
    }

    if (user.role === 'tutor') {
      setDb(p => ({
        ...p,
        tutors: p.tutors.map(t => t.username === user.username ? { ...t, password: newPwd, updatedAt: getSyncTimestamp() } : t)
      }));
    } else {
      setDb(p => ({
        ...p,
        users: p.users.map(u => u.username === user.username ? { ...u, password: newPwd, updatedAt: getSyncTimestamp() } : u)
      }));
    }

    setCurrentUser({ ...user, password: newPwd });

    // FIX #6: Force-push manual dihapus karena menggunakan snapshot db (stale closure)
    // yang bisa menimpa data terbaru dari perangkat lain.
    // Auto-sync useEffect akan mendeteksi perubahan setDb di atas dan menyinkronkan
    // ke GAS dalam ~2 detik secara aman dengan state React terbaru.

    showToast('Your password has been successfully updated.', 'success');
    setCurrentPwd(''); setNewPwd(''); setConfirmPwd('');
  };

  let profileDetails = null;
  if (user.role === 'student') {
    profileDetails = db.students.find(s => s.id === user.studentId);
  } else if (user.role === 'tutor') {
    profileDetails = db.tutors.find(t => t.username === user.username || t.name === user.name);
  }

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div>
        <h2 className="text-2xl font-bold text-white mb-2">{language === 'id' ? 'Profil & Pengaturan Saya' : 'My Profile & Settings'}</h2>
        <p className="text-gray-400 text-sm">{language === 'id' ? 'Lihat profil pribadi Anda dan kelola kredensial keamanan Anda.' : 'View your personal profile and manage your security credentials.'}</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card className="border border-[#00D4FF]/20 shadow-[0_0_30px_rgba(0,212,255,0.05)] h-max">
          <div className="flex items-center gap-3 mb-6 border-b border-gray-800 pb-4">
            <User className="text-[#00D4FF]" size={24} />
            <h3 className="text-lg font-bold text-white">{language === 'id' ? 'Informasi Profil' : 'Profile Information'}</h3>
          </div>
          <div className="space-y-4">
             <div>
                <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">{language === 'id' ? 'Nama Lengkap' : 'Full Name'}</p>
                <p className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800">{user.name}</p>
             </div>
             <div>
                <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">{language === 'id' ? 'Nama Pengguna' : 'Username'}</p>
                <p className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800">{user.username}</p>
             </div>
             <div>
                <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">{language === 'id' ? 'Peran / Status' : 'Role / Status'}</p>
                <div className="bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800 flex items-center gap-3">
                  <span className="text-white font-medium capitalize">{user.role}</span>
                  <Badge status={user.active || 'Active'} />
                </div>
             </div>
             
             {user.role === 'student' && profileDetails && (
               <>
                 <div>
                    <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">{language === 'id' ? 'Jenjang & Kelas' : 'Level & Class'}</p>
                    <p className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800">{profileDetails.level} - {profileDetails.class}</p>
                 </div>
                 {profileDetails.whatsapp && (
                   <div>
                      <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">WhatsApp</p>
                      {/* FIX #7: Konversi balik 628xxx → 08xxx untuk tampilan agar sesuai ekspektasi siswa */}
                      <p className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800">{(() => { const d = String(profileDetails.whatsapp || '').replace(/^'/, ''); return d.startsWith('628') ? '0' + d.slice(2) : d; })()}</p>
                   </div>
                 )}
               </>
             )}

             {user.role === 'tutor' && profileDetails && (
               <>
                 <div>
                    <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">{language === 'id' ? 'Sesi Mengajar' : 'Teaching Session'}</p>
                    <p className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800">{parseSessions(profileDetails.teachingSession).join(' · ') || '-'}</p>
                 </div>
                 {profileDetails.phone && (
                   <div>
                      <p className="text-xs text-gray-500 uppercase tracking-widest font-bold mb-1">WhatsApp</p>
                      <a href={`https://wa.me/${normalizeWhatsapp(profileDetails.phone)}`} target="_blank" rel="noopener noreferrer" className="text-white font-medium bg-[#0B0F19] px-4 py-2.5 rounded-lg border border-gray-800 flex items-center justify-between gap-2 hover:border-green-500/50 transition-colors">
                        {/* FIX #7: Konversi balik 628xxx → 08xxx untuk tampilan */}
                        <span>{(() => { const d = String(profileDetails.phone || '').replace(/^'/, ''); return d.startsWith('628') ? '0' + d.slice(2) : d; })()}</span>
                        <MessageCircle size={16} className="text-green-400" />
                      </a>
                   </div>
                 )}
               </>
             )}
          </div>
        </Card>

        <Card className="border border-[#00D4FF]/20 shadow-[0_0_30px_rgba(0,212,255,0.05)] h-max">
          <div className="flex items-center gap-3 mb-6 border-b border-gray-800 pb-4">
            <ShieldCheck className="text-[#00D4FF]" size={24} />
            <h3 className="text-lg font-bold text-white">{language === 'id' ? 'Ubah Kata Sandi' : 'Change Password'}</h3>
          </div>
          <form onSubmit={handleUpdate} className="space-y-5">
            <Input label={language === 'id' ? 'Kata Sandi Saat Ini' : 'Current Password'} type="password" value={currentPwd} onChange={setCurrentPwd} required placeholder={language === 'id' ? 'Masukkan kata sandi saat ini' : 'Enter your current password'} />
            <Input label={language === 'id' ? 'Kata Sandi Baru' : 'New Password'} type="password" value={newPwd} onChange={setNewPwd} required placeholder={language === 'id' ? 'Masukkan kata sandi baru yang aman' : 'Enter new secure password'} />
            <Input label={language === 'id' ? 'Konfirmasi Kata Sandi Baru' : 'Confirm New Password'} type="password" value={confirmPwd} onChange={setConfirmPwd} required placeholder={language === 'id' ? 'Ketik ulang kata sandi baru' : 'Re-type new password'} />
            <div className="pt-4 border-t border-gray-800">
              <Button type="submit" className="w-full text-lg py-3">{language === 'id' ? 'Perbarui Kredensial Keamanan' : 'Update Security Credentials'}</Button>
              <p className="text-center text-xs text-gray-500 mt-4">{language === 'id' ? 'Catatan: Kata sandi disimpan secara aman di sistem.' : 'Note: Passwords are saved in the system securely.'}</p>
            </div>
          </form>
        </Card>
      </div>
    </div>
  );
}

// MODULES KHUSUS STUDENT (READ-ONLY)

function StudentReadOnlyAttendanceModule({ db, user, language = 'en' }) {
  const [filterMonth, setFilterMonth] = useState<number | string>(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());

  const myAtt = db.studentAttendance.filter(a => {
    if (a.studentId !== user.studentId) return false;
    if (filterMonth !== 'All') {
       const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
       return a.date.startsWith(prefix);
    }
    return a.date.startsWith(String(filterYear));
  }).reverse();
  
  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Rekam Kehadiran Saya' : 'My Attendance Record'}</h2><p className="text-gray-400 text-sm">{language === 'id' ? 'Lihat riwayat kehadiran Anda.' : 'View your attendance history.'}</p></div>
      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-wrap gap-3 justify-between items-center">
           <div className="flex gap-2 w-full sm:w-auto">
              <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                <option value="All">{language === 'id' ? 'Semua Bulan' : 'All Months'}</option>
                {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
              </select>
              <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
           </div>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr><th className="p-4 text-center w-12 text-gray-400">No.</th><th className="p-4 text-center">{language === 'id' ? 'Tanggal' : 'Date'}</th><th className="p-4 text-center">{language === 'id' ? 'Sesi' : 'Session Group'}</th><th className="p-4 text-center">Status</th></tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {myAtt.length > 0 ? myAtt.map((a, index) => (
              <tr key={a.id} className="hover:bg-[#0B0F19]">
                <td className="p-4 text-center"><span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF]">{String(index + 1).padStart(2, '0')}</span></td>
                <td className="p-4 text-center font-medium text-white">{a.date}</td>
                <td className="p-4 text-center text-gray-400">{a.sessionGroup}</td>
                <td className="p-4 text-center"><Badge status={a.status} /></td>
              </tr>
            )) : <tr><td colSpan={4}><EmptyState icon={UserCheck} title={language === 'id' ? 'Belum ada rekam kehadiran' : 'No attendance recorded yet'} description={language === 'id' ? 'Data kehadiran akan muncul di sini.' : 'Attendance records will appear here.'} /></td></tr>}
          </tbody>
        </table>
        </div>
      </Card>
    </div>
  )
}

function StudentReadOnlyJournalsModule({ db, user, language = 'en' }) {
  const [filterMonth, setFilterMonth] = useState<number | string>(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());

  const student = db.students.find(s => s.id === user.studentId);
  const myGroup = student ? getStudentSession(student) : '';
  // FIX F1: Normalisasi sessionGroup — beberapa jurnal mungkin disimpan dengan
  // sched.name (nama bebas dari kalender) bukan nilai SESSIONS[] yang baku.
  // Fallback: jika j.sessionGroup mengandung bagian dari myGroup (atau sebaliknya),
  // tetap tampilkan. Ini mencegah siswa kehilangan jurnal karena typo nama sesi.
  const sessionMatches = (jGroup, sGroup) => {
    if (!jGroup || !sGroup) return false;
    if (jGroup === sGroup) return true;
    // Fuzzy: salah satu mengandung yang lain (case-insensitive)
    const a = jGroup.toLowerCase(), b = sGroup.toLowerCase();
    return a.includes(b) || b.includes(a);
  };
  const myJournals = db.journals.filter(j => {
     if (!sessionMatches(j.sessionGroup, myGroup)) return false;
     if (filterMonth !== 'All') {
       const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
       return j.date.startsWith(prefix);
     }
     return j.date.startsWith(String(filterYear));
  }).reverse();

  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Materi Belajar Saya' : 'My Learning Materials'}</h2><p className="text-gray-400 text-sm">{language === 'id' ? 'Akses materi dan sumber belajar Anda.' : 'Access your learning materials and resources.'}</p></div>
      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-wrap gap-3 justify-between items-center">
           <div className="flex gap-2 w-full sm:w-auto">
              <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                <option value="All">{language === 'id' ? 'Semua Bulan' : 'All Months'}</option>
                {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
              </select>
              <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
           </div>
        </div>
        <div className="p-4 sm:p-5 space-y-4 bg-[#151B26]">
        {myJournals.length > 0 ? myJournals.map((j, idx) => (
          <div key={j.id} className="bg-[#0B0F19] border border-[#00D4FF]/20 rounded-xl shadow-md p-5">
            <div className="flex justify-between items-start mb-2">
              <div className="flex items-start gap-3">
                 <span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF] shrink-0 mt-0.5">{String(idx + 1).padStart(2, '0')}</span>
                 <div>
                    <h3 className="text-lg font-bold text-white">{j.topic}</h3>
                    <p className="text-[#00D4FF] text-sm">Tutor: {j.tutorName}</p>
                 </div>
              </div>
              <span className="text-xs text-gray-400 shrink-0 ml-2">{j.date}</span>
            </div>
            <div className="mt-4 pt-4 border-t border-gray-800">
              <p className="text-sm text-gray-300"><strong>{language === 'id' ? 'Activities:' : 'Activities:'}</strong> {j.activities}</p>
              {j.followUp && <p className="text-sm text-gray-400 mt-2"><strong>{language === 'id' ? 'Tindak Lanjut:' : 'Follow Up:'}</strong> {j.followUp}</p>}
            </div>
          </div>
        )) : <div className="p-8 text-center text-gray-500 bg-[#151B26] rounded-xl">{language === 'id' ? 'Tidak ada materi tersedia untuk sesi Anda.' : 'No materials available for your session.'}</div>}
        </div>
      </Card>
    </div>
  )
}

function StudentReadOnlyAssessmentModule({ db, user, language = 'en' }) {
  const [filterMonth, setFilterMonth] = useState<number | string>('All');
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());

  const myAssessments = db.assessments
    .filter(a => {
       if (a.studentId !== user.studentId) return false;
       if (Number(a.year) !== filterYear) return false;
       if (filterMonth !== 'All' && Number(a.month) !== filterMonth) return false;
       return true;
    })
    .sort((a, b) => {
       if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
       return Number(b.month) - Number(a.month);
    });

  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Hasil Penilaian Saya' : 'My Assessment Results'}</h2><p className="text-gray-400 text-sm">{language === 'id' ? 'Lihat nilai dan hasil penilaian Anda.' : 'View your grades and assessment results.'}</p></div>
      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-wrap gap-3 justify-between items-center">
           <div className="flex gap-2 w-full sm:w-auto">
              <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                <option value="All">{language === 'id' ? 'Semua Bulan' : 'All Months'}</option>
                {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
              </select>
              <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} placeholder={language === 'id' ? "Tahun" : "Year"} />
           </div>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr>
              <th className="p-4 text-center w-12 text-gray-400">No.</th>
              <th className="p-4 text-center">{language === 'id' ? 'Periode' : 'Period'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Rata-rata' : 'Average Score'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Nilai Akhir' : 'Final Grade'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Detail' : 'Details'}</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {myAssessments.length > 0 ? myAssessments.map((a, index) => (
              <tr key={a.id} className="hover:bg-[#0B0F19]">
                <td className="p-4 text-center"><span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF]">{String(index + 1).padStart(2, '0')}</span></td>
                <td className="p-4 text-center font-medium text-white">{MONTHS[parseInt(a.month)-1]} {a.year}</td>
                <td className="p-4 text-center font-bold text-[#00D4FF] text-lg">{a.average}</td>
                <td className="p-4 text-center font-bold text-white text-lg">{a.grade}</td>
                <td className="p-4 text-center text-gray-400 text-xs">
                   {Object.entries(a.scores || {}).map(([sub, score]) => (
                      <span key={sub} className="block">{`${sub}: ${score}`}</span>
                   ))}
                </td>
              </tr>
            )) : <tr><td colSpan={5}><EmptyState icon={Award} title={language === 'id' ? 'Belum ada penilaian' : 'No assessments recorded yet'} description={language === 'id' ? 'Data penilaian akan muncul di sini.' : 'Assessment records will appear here.'} /></td></tr>}
          </tbody>
        </table>
        </div>
      </Card>
    </div>
  )
}

function StudentReadOnlyPaymentModule({ db, user, downloadPNG, handleShareImage, language = 'en', showToast }) {
  const [filterMonth, setFilterMonth] = useState<number | string>('All');
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  const [selectedInvoice, setSelectedInvoice] = useState(null);

  const myPayments = db.payments
    .filter(p => {
       if (p.studentId !== user.studentId) return false;
       if (Number(p.year) !== filterYear) return false;
       if (filterMonth !== 'All' && Number(p.month) !== filterMonth) return false;
       return true;
    })
    .sort((a, b) => {
       if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
       return Number(b.month) - Number(a.month);
    });

  const localPrintPayment = () => {
    if (!selectedInvoice) return;
    const originalTitle = document.title;
    const safeName = (selectedInvoice.studentName || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_');
    document.title = `${safeName}_payment`;
    window.print();
    setTimeout(() => { document.title = originalTitle; }, 1000);
  };

  if (selectedInvoice) return (
    <div className="fixed inset-0 z-[100] bg-slate-50/95 backdrop-blur-md overflow-y-auto print:bg-white print:static print:block print:z-auto custom-scrollbar font-sans text-slate-900">
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[80%] h-[400px] bg-gradient-to-b from-blue-200/40 to-transparent blur-3xl pointer-events-none print:hidden" />
      
      <div className="w-full max-w-2xl mx-auto mt-6 mb-4 px-4 flex justify-between items-center relative z-10 print:hidden">
        <button className="flex items-center gap-2 text-slate-500 hover:text-slate-800 transition-colors font-medium text-sm bg-white/50 px-4 py-2 rounded-full border border-slate-200/50 shadow-sm" onClick={() => setSelectedInvoice(null)}>
          <ArrowLeft size={16} /> Back
        </button>
        <div className="flex gap-2">
          <button onClick={localPrintPayment} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Print PDF"><Printer size={16}/></button>
          <button onClick={() => {
             const safeName = (selectedInvoice.studentName || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_');
             downloadPNG('receipt-print', `${safeName}_payment`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Download PNG"><Download size={16}/></button>
          <button onClick={() => {
             const safeName = (selectedInvoice.studentName || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_');
             handleShareImage('receipt-print', `${safeName}_payment`, `Receipt for ${selectedInvoice.studentName || 'student'}`);
          }} className="p-2.5 bg-white rounded-full text-blue-600 shadow-sm border border-slate-200/50 hover:bg-blue-50 transition-colors" title="Share"><Share2 size={16}/></button>
        </div>
      </div>

      <div id="receipt-print" className="w-full max-w-2xl mx-auto bg-white rounded-none shadow-2xl overflow-hidden relative z-10 mb-12 print:m-0 print:shadow-none print:max-w-full print:rounded-none">
        <div className="bg-[#1A56DB] text-white p-6 sm:p-8 flex justify-between items-start">
          <div>
            <h1 className="text-2xl sm:text-3xl font-black tracking-tight">PAYMENT RECEIPT</h1>
            <p className="text-blue-200 font-mono mt-1 text-sm">{selectedInvoice.id}</p>
          </div>
          <div className="text-right">
            <p className="text-lg font-bold text-white leading-tight">English Club Gresik</p>
            <p className="text-blue-200 text-xs mt-1">Academic Suite</p>
          </div>
        </div>

        <div className="p-6 sm:p-8 pb-4">
          <div className="text-center mb-8 border-b border-slate-200 pb-8">
            <p className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">
              Tuition Fee {MONTHS[parseInt(selectedInvoice.month) - 1]} {selectedInvoice.year}
            </p>
            <p className="text-5xl font-black text-slate-900 tracking-tight">
              Rp {Number(selectedInvoice.amount).toLocaleString('id-ID')}
            </p>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">STUDENT INFO</p>
            <div className="bg-slate-50 border border-slate-200 p-5 shadow-sm">
              <p className="text-xl font-bold text-slate-800">{selectedInvoice.studentName}</p>
              <p className="text-sm font-medium text-slate-500 mt-1">{selectedInvoice.class} · {selectedInvoice.sessionGroup}</p>
            </div>
          </div>

          <div className="mb-6">
            <p className="text-[11px] text-[#1A56DB] font-bold uppercase tracking-wider mb-2 px-2">TRANSACTION DETAILS</p>
            <div className="space-y-0 text-sm">
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Date & Time</span>
                <span className="font-semibold text-slate-800">
                  {/* FIX #5: Gunakan safeDateDisplay agar konsisten dan bebas UTC-shift */}
                  {safeDateDisplay(selectedInvoice.date, 'en-US', { month: 'long', day: 'numeric', year: 'numeric' })} {selectedInvoice.time ? `• ${selectedInvoice.time}` : ''}
                </span>
              </div>
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Method</span>
                <span className="font-semibold text-slate-800">{selectedInvoice.method || 'Cash'}</span>
              </div>
              <div className="flex justify-between items-center border-b border-slate-100 py-3 px-2">
                <span className="font-medium text-slate-600">Receipt No.</span>
                <span className="font-mono font-semibold text-slate-800">{selectedInvoice.id}</span>
              </div>
              <div className="flex justify-between items-center py-3 px-2">
                <span className="font-medium text-slate-600">Status</span>
                <span className="text-lg font-black text-green-700">PAID</span>
              </div>
            </div>
          </div>

          <div className="text-center mt-12 mb-4">
            <p className="text-sm font-medium text-slate-600 italic">Thank you for your payment.</p>
            <p className="text-xs font-bold text-slate-500 mt-1">— Akhmad Akmal Rifqi</p>
          </div>
        </div>

          <div className="bg-[#1A56DB] p-6 sm:p-8 text-center">
            <p className="text-xs text-blue-100 font-medium mb-1">English Club Gresik • WA: 0897-327-11-12</p>
            <p className="text-[11px] text-blue-200 mb-3">Taman Anggrek Blok AB 05, Kedanyang, Kebomas, Gresik</p>
            <div className="border-t border-blue-400/30 pt-3">
              <p className="text-[11px] text-blue-200">
                This document serves as an official payment receipt
              </p>
            </div>
          </div>
        </div>
      </div>
    );

  return (
    <div className="space-y-6">
      <div><h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Riwayat Pembayaran Saya' : 'My Payment History'}</h2><p className="text-gray-400 text-sm">{language === 'id' ? 'Lihat riwayat pembayaran Anda.' : 'View your payment history.'}</p></div>

      <div className="bg-gradient-to-r from-orange-500/20 to-orange-600/10 border border-orange-500/30 p-5 rounded-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-5 shadow-lg relative overflow-hidden animation-fade-in">
          <div className="absolute -right-10 -top-10 w-40 h-40 bg-orange-500/10 blur-3xl rounded-full"></div>
          <div className="relative z-10 flex items-center gap-4">
              <div className="w-14 h-14 bg-white rounded-xl flex items-center justify-center font-black text-orange-600 text-xl shadow-md shrink-0 border-2 border-orange-500/20">
                  BNI
              </div>
              <div>
                  <p className="text-xs text-orange-400 font-bold uppercase tracking-widest mb-0.5">{language === 'id' ? 'Rekening Pembayaran Resmi' : 'Official Payment Account'}</p>
                  <p className="text-lg sm:text-xl font-bold text-white leading-tight">Akhmad Akmal Rifqi</p>
                  <p className="text-xs sm:text-sm text-gray-300 font-medium mt-0.5">Bank Negara Indonesia (BNI)</p>
              </div>
          </div>
          <div className="relative z-10 flex items-center justify-between w-full md:w-auto gap-4 bg-[#0B0F19]/80 backdrop-blur-sm px-5 py-3 rounded-xl border border-gray-700 shadow-inner">
              <span className="text-2xl sm:text-3xl font-mono font-black text-[#00D4FF] tracking-wider drop-shadow-md">0951837774</span>
              <button
                  onClick={() => {
                      // BUGFIX: Tambahkan optional chaining agar tidak crash di browser yang tidak support clipboard API
                      navigator.clipboard?.writeText('0951837774');
                      showToast(language === 'id' ? 'Nomor rekening BNI disalin!' : 'BNI Account number copied!');
                  }}
                  className="bg-[#00D4FF]/10 hover:bg-[#00D4FF]/20 text-[#00D4FF] p-2.5 rounded-lg transition-all active:scale-95 shadow-sm border border-[#00D4FF]/20"
                  title={language === 'id' ? 'Salin Rekening' : 'Copy Account Number'}
              >
                  <Copy size={20} />
              </button>
          </div>
      </div>

      <Card className="p-0 flex flex-col">
        <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-wrap gap-3 justify-between items-center">
           <div className="flex gap-2 w-full sm:w-auto">
              <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                <option value="All">{language === 'id' ? 'Semua Bulan' : 'All Months'}</option>
                {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
              </select>
              <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} placeholder={language === 'id' ? "Tahun" : "Year"} />
           </div>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
            <tr>
              <th className="p-4 text-center w-12 text-gray-400">No.</th>
              <th className="p-4 text-center">{language === 'id' ? 'ID Tagihan' : 'Invoice ID'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Periode' : 'Period'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Jumlah Bayar' : 'Amount Paid'}</th>
              <th className="p-4 text-center">{language === 'id' ? 'Tanggal' : 'Date'}</th>
              <th className="p-4 text-center">Status</th>
              <th className="p-4 text-center">{language === 'id' ? 'Kuitansi' : 'Receipt'}</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-800">
            {myPayments.length > 0 ? myPayments.map((p, index) => (
              <tr key={p.id} className="hover:bg-[#0B0F19]">
                <td className="p-4 text-center"><span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF]">{String(index + 1).padStart(2, '0')}</span></td>
                <td className="p-4 text-center font-mono text-gray-500 text-xs">{p.id}</td>
                <td className="p-4 text-center font-medium text-white">{MONTHS[parseInt(p.month)-1]} {p.year}</td>
                <td className="p-4 text-center font-bold text-green-400">Rp {Number(p.amount).toLocaleString()}</td>
                <td className="p-4 text-center text-gray-400">{p.date}</td>
                <td className="p-4 text-center"><Badge status={p.status} /></td>
                <td className="p-4 text-center">
                  {p.status === 'Paid' ? (
                     <Button variant="ghost" className="text-xs px-3 py-1 h-8 mx-auto flex items-center gap-2" onClick={() => setSelectedInvoice(p)}><FileText size={14} /> {language === 'id' ? 'Unduh' : 'Download'}</Button>
                  ) : (
                     <span className="text-gray-500 text-xs">-</span>
                  )}
                </td>
              </tr>
            )) : <tr><td colSpan={7}><EmptyState icon={DollarSign} title={language === 'id' ? 'Tidak ada riwayat pembayaran' : 'No payment records found'} description={language === 'id' ? 'Riwayat pembayaran akan muncul di sini.' : 'Payment history will appear here.'} /></td></tr>}
          </tbody>
        </table>
        </div>
      </Card>
    </div>
  )
}

// ADD START
function StudentReadOnlyReportModule({ db, user, downloadPNG, handleShareImage, language = 'en' }) {
  const [reportMonth, setReportMonth] = useState(new Date().getMonth() + 1);
  const [reportYear, setReportYear] = useState(new Date().getFullYear());

  const student = db.students.find(s => s.id === user.studentId);
  if (!student) return <div className="p-8 text-center text-red-500">Student profile not found.</div>;

  const reportPrefix = `${reportYear}-${String(reportMonth).padStart(2, '0')}`;
  
  const att = db.studentAttendance.filter((a) => a.studentId === student.id && a.date.startsWith(reportPrefix));
  const presentCount = att.filter((a) => a.status === 'Present').length;
  const sickCount = att.filter((a) => a.status === 'Sick').length;
  const excusedCount = att.filter((a) => a.status === 'Excused').length;
  const absentCount = att.filter((a) => a.status === 'Absent').length;
  const attRate = att.length ? Math.round((presentCount / att.length) * 100) : 0;
  
  const assessments = db.assessments.filter((a) => a.studentId === student.id && Number(a.month) === reportMonth && Number(a.year) === reportYear);
  const avgScore = assessments.length ? Math.round(assessments.reduce((sum, a) => sum + a.average, 0) / assessments.length) : 0;
  
  // BUGFIX: Gunakan fuzzy match agar jurnal tidak kosong di report card karena perbedaan format nama sesi
  const _reportStudentSession = getStudentSession(student);
  const journals = db.journals.filter((j) => {
    if (!j.date.startsWith(reportPrefix)) return false;
    const jg = (j.sessionGroup || '').toLowerCase();
    const sg = (_reportStudentSession || '').toLowerCase();
    return jg === sg || jg.includes(sg) || sg.includes(jg);
  });

  const studentSession = getStudentSession(student);
  const reportSubjects = studentSession === SESSIONS[0]
    ? ['Reading', 'Writing', 'Math', 'English']
    : ['Speaking', 'Writing', 'Reading', 'Listening'];

  const isKindergarten = student.level === 'Kindergarten' || ['PAUD', 'TK A', 'TK B'].includes(student.class);
  // Fix #9: sort agar [0] selalu record terbaru (handle kemungkinan duplikat)
  const sortedReportAssessments = [...assessments].sort((a, b) => {
    if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
    return Number(b.month) - Number(a.month);
  });
  const latestAss = sortedReportAssessments.length > 0 ? sortedReportAssessments[0] : null;
  const scores = latestAss && latestAss.scores ? latestAss.scores : {};
  const getScore = (subject) => Number(scores[subject]) || 0;

  const skillProgress = isKindergarten ? [
     { label: 'Reading', value: getScore('Reading') },
     { label: 'Writing', value: getScore('Writing') },
     { label: 'Math', value: getScore('Math') },
     { label: 'English', value: getScore('English') }
  ] : [
     { label: 'Speaking', value: getScore('Speaking') },
     { label: 'Writing', value: getScore('Writing') },
     { label: 'Reading', value: getScore('Reading') },
     { label: 'Listening', value: getScore('Listening') }
  ];

  const hasAssessments = assessments.length > 0;
  let finalCommentDisplay = "";

  if (!hasAssessments) {
     finalCommentDisplay = language === 'id' 
         ? "Komentar akademik belum tersedia karena penilaian untuk periode ini belum diselesaikan oleh tutor." 
         : "Academic comments are not available yet as assessments for this period have not been completed by the tutor.";
  } else {
     const autoGeneratedComment = generateAutoComment(student, attRate, avgScore, assessments);
     finalCommentDisplay = student.teacherComment || autoGeneratedComment;
  }

  const handlePrintStudentReport = () => {
    const originalTitle = document.title;
    const safeName = (student?.name || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_');
    document.title = `${safeName}_report_${MONTHS[reportMonth-1]}_${reportYear}`;
    window.print();
    setTimeout(() => { document.title = originalTitle; }, 1000);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center bg-[#0A0E17] p-4 rounded-xl border border-gray-800 shadow-sm gap-4 print-hidden">
        <div>
           <h2 className="text-xl font-bold text-white">{language === 'id' ? 'Rapor Akademik Saya' : 'My Academic Report'}</h2>
           <p className="text-sm text-gray-400">{language === 'id' ? 'Lihat dan unduh rapor perkembangan bulanan Anda' : 'View and download your monthly progress report'}</p>
        </div>
        <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto">
          <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF]" value={reportMonth} onChange={e => setReportMonth(Number(e.target.value))}>
            {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
          </select>
          <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={reportYear} onChange={e => setReportYear(Number(e.target.value))} />
          <div className="flex gap-2 w-full sm:w-auto mt-2 sm:mt-0">
             <Button onClick={handlePrintStudentReport} icon={Printer} className="flex-1 sm:flex-none">{language === 'id' ? 'Cetak PDF' : 'Print PDF'}</Button>
             <Button onClick={() => downloadPNG('report-print', `${(student?.name || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_')}_report`)} variant="secondary" icon={Download} title="Download Image" className="px-3" />
             <Button onClick={() => handleShareImage('report-print', `${(student?.name || 'student').replace(/[\s/\\?%*:|"<>-]/g, '_')}_report`, `Academic Report for ${student?.name || 'student'}`)} variant="secondary" icon={Share2} title="Share" className="px-3" />
          </div>
        </div>
      </div>

      {/* Premium PDF Layout - Read Only */}
      <div className="w-full animation-fade-in relative bg-white rounded-xl shadow-2xl p-8 border-2 border-[#1A56DB] print:p-0 print:border-2 print:border-[#1A56DB] print:shadow-none print:w-full font-sans max-w-4xl mx-auto text-[11px] text-slate-900" id="report-print">
        
        {/* Subtle Watermark */}
        <div className="fixed inset-0 flex items-center justify-center opacity-[0.02] z-0 pointer-events-none hidden print:flex">
          <img src={LOGO_URL} className="w-[400px] h-auto grayscale" alt="watermark" />
        </div>

        <div className="relative z-10">
          
          {/* 1. PREMIUM HEADER */}
          <div className="flex justify-between items-center border-b border-slate-200 pb-4 mb-5 print:mt-4">
            <div className="flex items-center gap-5">
              <img src={LOGO_URL} className="h-16 object-contain" alt="Logo" />
            </div>
            <div className="text-right">
              <h1 className="text-[26px] font-bold text-slate-900 tracking-tight uppercase leading-none">Monthly Academic Progress Report</h1>
              <h2 className="text-[11px] font-medium text-slate-500 mt-2 uppercase tracking-[0.8px]">English Club Gresik • Report Period: {MONTHS[reportMonth - 1]} {reportYear}</h2>
            </div>
          </div>

          {/* 2. STUDENT PROFILE & DASHBOARD SUMMARY CARDS */}
          <div className="grid grid-cols-1 md:grid-cols-12 print:grid-cols-12 gap-3 mb-4 break-inside-avoid">
            {/* Profile Card */}
            <div className="md:col-span-5 print:col-span-5 bg-[#F8FAFC] border border-slate-200 rounded-[10px] p-4">
              <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Student Profile</p>
              <h4 className="text-lg font-bold text-slate-900 mb-0.5">{student.name}</h4>
              <p className="text-[11px] font-mono text-slate-500 mb-3">{student.id}</p>
              <div className="flex gap-6 text-[11px]">
                <div><span className="text-slate-400 font-medium text-[11px] uppercase tracking-[0.5px] block mb-0.5">Class Level</span><span className="font-semibold text-slate-900">{student.level} - {student.class}</span></div>
                <div><span className="text-slate-400 font-medium text-[11px] uppercase tracking-[0.5px] block mb-0.5">Session</span><span className="font-semibold text-slate-900">{studentSession}</span></div>
              </div>
            </div>

            {/* Skill Radar Chart */}
            <div className="md:col-span-3 print:col-span-3 bg-white border border-slate-200 rounded-[10px] p-3 shadow-sm flex flex-col items-center justify-center relative">
               <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px] absolute top-3 left-3">Skill Analytics</p>
               <div className="w-full max-w-[110px] aspect-square mt-3">
                   <RadarChart data={skillProgress} theme="light" />
               </div>
            </div>

            {/* Dashboard Summary Cards */}
            <div className="md:col-span-4 print:col-span-4 flex flex-col gap-2">
               {/* Premium Attendance Card */}
               <div className="flex-1 bg-white border border-slate-200 rounded-[10px] p-2.5 shadow-sm flex flex-col justify-between">
                  <div className="flex justify-between items-center mb-1">
                     <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px]">Attendance</p>
                     <p className="text-xl font-black text-slate-900 leading-none">{(att.length > 0) ? `${attRate}%` : '—'}</p>
                  </div>
                  {att.length > 0 && (
                     <div className="bg-[#F8FAFC] border border-slate-100 rounded-lg p-1.5 grid grid-cols-2 gap-x-2 gap-y-0.5 text-[8px]">
                        <div className="flex justify-between"><span className="text-slate-500">Present</span><span className="font-bold">{presentCount}</span></div>
                        <div className="flex justify-between"><span className="text-slate-500">Sick</span><span className="font-bold">{sickCount}</span></div>
                        <div className="flex justify-between"><span className="text-slate-500">Excused</span><span className="font-bold">{excusedCount}</span></div>
                        <div className="flex justify-between"><span className="text-slate-500">Absent</span><span className="font-bold">{absentCount}</span></div>
                     </div>
                  )}
               </div>

               {/* Premium Avg Score Card */}
               <div className="flex-1 bg-white border border-slate-200 rounded-[10px] p-2.5 shadow-sm flex flex-col justify-between">
                  <div className="flex justify-between items-center mb-1">
                     <p className="text-[11px] text-slate-400 font-bold uppercase tracking-[0.8px]">Avg Score</p>
                     <p className="text-xl font-black text-slate-900 leading-none">{assessments.length > 0 ? avgScore : '—'}</p>
                  </div>
                  {assessments.length > 0 && (
                     <div className="bg-[#F8FAFC] border border-slate-100 rounded-lg p-1.5 flex items-center justify-between h-[30px]">
                        <span className="text-[11px] text-slate-500 font-medium uppercase tracking-wide">Final Grade</span>
                        {/* Fix #9: sortedReportAssessments[0] — record terbaru */}
                        <span className="font-black text-base text-blue-600 leading-none">{sortedReportAssessments[0]?.grade || '—'}</span>
                     </div>
                  )}
               </div>
            </div>
          </div>

          {/* 3. ACADEMIC ADVISOR COMMENTS */}
          <div className="mb-5 break-inside-avoid">
             <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Academic Advisor Comments</p>
             <div className="bg-[#F8FAFC] border-l-4 border-blue-600 p-[16px] rounded-[10px]">
                <p className={`text-[11px] leading-relaxed text-slate-800 ${(!assessments.length) ? 'italic text-slate-500' : 'font-medium'}`}>
                   {assessments.length ? `"${finalCommentDisplay}"` : "No comments recorded for this period."}
                </p>
             </div>
          </div>

          {/* 4. ASSESSMENT TABLE */}
          <div className="mb-5 break-inside-avoid">
            <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Monthly Assessment Grades</p>
            {(assessments.length > 0) ? (
                <div className="w-full border border-slate-200 rounded-[10px] overflow-hidden">
                    <table className="w-full text-[11px] text-left">
                      <thead className="bg-[#EFF6FF] text-[#1E3A8A]">
                        <tr>
                          <th className="py-2.5 px-4 font-semibold w-24">Period</th>
                          {reportSubjects.map(sub => (
                            <th key={sub} className="py-2.5 px-4 text-center font-semibold">{sub}</th>
                          ))}
                          <th className="py-2.5 px-4 text-center font-semibold w-20">Average</th>
                          <th className="py-2.5 px-4 text-center font-semibold w-16">Grade</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {assessments.map((a, idx) => (
                          <tr key={a.id} className="bg-white even:bg-[#FAFAFA]">
                            <td className="py-2.5 px-4 font-semibold text-slate-900">{MONTHS[parseInt(a.month)-1].substring(0,3)} '{String(a.year).slice(2)}</td>
                            {reportSubjects.map(sub => (
                              <td key={sub} className="py-2.5 px-4 text-center text-slate-700">{a.scores?.[sub] || '—'}</td>
                            ))}
                            <td className="py-2.5 px-4 text-center font-bold text-slate-900">{a.average || '—'}</td>
                            <td className="py-2.5 px-4 text-center font-bold text-blue-600">{a.grade || '—'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                </div>
            ) : (
                <div className="bg-[#F8FAFC] border border-slate-200 rounded-[10px] py-3 px-4 flex items-center gap-2 text-slate-500 w-max">
                    <FileText size={14} className="text-slate-400" />
                    <span className="italic text-[11px]">No assessments recorded for this period.</span>
                </div>
            )}
          </div>

          {/* 5. SESSION DETAIL & ATTENDANCE LOG */}
          <div className="mb-6 break-inside-avoid">
            <p className="text-[11px] font-bold text-slate-400 uppercase tracking-[0.8px] mb-1.5">Session Detail & Attendance Log</p>
            {(() => {
              // FIX: Anchor ke attendance siswa saja (bukan union dengan journal).
              const sortedAtt = [...att].sort((a, b) =>
                new Date(a.date).getTime() - new Date(b.date).getTime()
              );

              const combinedSessions = sortedAtt.map((attEntry, idx) => {
                  const journalEntry = journals.find(j => j.date === attEntry.date && j.scheduleId === attEntry.scheduleId)
                    || journals.find(j => j.date === attEntry.date);
                  const status = attEntry.status || "—";
                  const sessionNum = idx + 1;
                  const streak = (() => {
                    let s = 0;
                    for (let i = idx; i >= 0; i--) {
                      if (sortedAtt[i].status === 'Present') s++; else break;
                    }
                    return s;
                  })();
                  const autoNote = (() => {
                    if (status === 'Present') {
                      if (sessionNum === 1) return "First session — student attended and engaged well.";
                      if (streak >= 5) return `Excellent consistency — ${streak} sessions in a row with full attendance.`;
                      if (streak >= 3) return `Good streak — attended ${streak} consecutive sessions.`;
                      const notes = [
                        "Attended and followed lesson activities.",
                        "Active participation noted during the session.",
                        "Session completed; student was present and attentive.",
                        "Student followed all learning materials for this session.",
                        "Good attendance — lesson was conducted as scheduled.",
                      ];
                      return notes[sessionNum % notes.length];
                    }
                    if (status === 'Sick') return "Student was absent due to illness. Material may need a brief review next session.";
                    if (status === 'Excused') return "Absence was excused. Student is advised to review session material independently.";
                    if (status === 'Absent') return "Student was absent without notice. Follow-up with parent/guardian recommended.";
                    return "Attendance status unrecorded for this session.";
                  })();
                  return {
                      date: attEntry.date,
                      topic: journalEntry?.topic || "No lesson record entered",
                      status,
                      note: (journalEntry?.notes && journalEntry.notes.trim() !== "") ? journalEntry.notes : autoNote
                  };
              });

              return combinedSessions.length > 0 ? (
                <div className="w-full border border-slate-200 rounded-[10px] overflow-hidden">
                    <table className="w-full text-[11px] text-left">
                      <thead className="bg-[#EFF6FF] text-[#1E3A8A]">
                        <tr>
                          <th className="py-2.5 px-4 font-semibold w-10 text-center">No</th>
                          <th className="py-2.5 px-4 font-semibold w-24">Date</th>
                          <th className="py-2.5 px-4 font-semibold">Material / Topic</th>
                          <th className="py-2.5 px-4 font-semibold text-center w-24">Status</th>
                          <th className="py-2.5 px-4 font-semibold text-right hidden sm:table-cell print:table-cell">Note</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {combinedSessions.map((session, index) => (
                          <tr key={session.date} className="bg-white even:bg-[#FAFAFA]">
                            <td className="py-2.5 px-4 text-center text-slate-500">{index + 1}</td>
                            <td className="py-2.5 px-4 whitespace-nowrap font-semibold text-slate-900">{new Date(session.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}</td>
                            <td className={`py-2.5 px-4 ${session.topic === 'No lesson record entered' ? 'text-slate-400 italic' : 'text-slate-700'}`}>{session.topic}</td>
                            <td className="py-2.5 px-4 text-center">
                               <span className="font-bold text-[11px] tracking-wider uppercase">
                                  {session.status !== '—' ? session.status : '—'}
                               </span>
                            </td>
                            <td className="py-2.5 px-4 text-right text-slate-400 italic text-[11px] hidden sm:table-cell print:table-cell truncate max-w-[180px]">{session.note}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                </div>
              ) : (
                <div className="bg-[#F8FAFC] border border-slate-200 rounded-[10px] py-3 px-4 flex items-center gap-2 text-slate-500 w-max">
                    <CalendarIcon size={14} className="text-slate-400" />
                    <span className="italic text-[11px]">No attendance records available for this period.</span>
                </div>
              );
            })()}
          </div>

          {/* 6. SIGNATURE SECTION */}
          <div className="flex justify-between items-end break-inside-avoid text-[11px] signature-section mb-6">
             <div></div>
             <div className="text-right w-48">
               <p className="text-slate-500 mb-8 font-medium">Gresik, {new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
               <div className="border-t border-slate-300 pt-1.5 inline-block w-full">
                   <p className="font-bold text-slate-900 tracking-wide">Akhmad Akmal Rifqi</p>
                   <p className="text-slate-500">Academic Advisor</p>
               </div>
               <p className="text-[11px] font-medium text-slate-400 mt-1 flex items-center justify-end gap-1"><CheckCircle2 size={10}/> Verified by ECG Academic Suite</p>
             </div>
          </div>

          {/* 7. PREMIUM FOOTER */}
          <div className="border-t border-slate-200 pt-3 text-[11px] text-slate-400 flex justify-between items-center hidden print:flex uppercase tracking-wide">
             <span>English Club Gresik Premium Report</span>
             <span>Generated automatically</span>
             <span>www.englishclub.my.id</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// --- SINTESIS WEB AUDIO API ---
let audioCtx: AudioContext | null = null;

const initAudioContext = () => {
  if (!audioCtx) {
    const AudioContextClass = window.AudioContext || (window as any).webkitAudioContext;
    if (AudioContextClass) {
       audioCtx = new AudioContextClass();
    }
  }
  if (audioCtx && audioCtx.state === 'suspended') {
    audioCtx.resume();
  }
  return audioCtx;
};

const playFeedbackSound = (type: 'success' | 'error') => {
  try {
    const ctx = initAudioContext();
    if (!ctx) return;
    
    const osc = ctx.createOscillator();
    const gainNode = ctx.createGain();

    osc.connect(gainNode);
    gainNode.connect(ctx.destination);

    const now = ctx.currentTime;

    if (type === 'success') {
      osc.type = 'sine';
      osc.frequency.setValueAtTime(523.25, now);
      osc.frequency.exponentialRampToValueAtTime(1046.50, now + 0.1);
      gainNode.gain.setValueAtTime(0, now);
      gainNode.gain.linearRampToValueAtTime(0.5, now + 0.02);
      gainNode.gain.exponentialRampToValueAtTime(0.01, now + 0.5);
      osc.start(now);
      osc.stop(now + 0.5);
    } else {
      osc.type = 'triangle';
      osc.frequency.setValueAtTime(300, now);
      osc.frequency.setValueAtTime(300, now + 0.1);
      osc.frequency.exponentialRampToValueAtTime(150, now + 0.15);
      gainNode.gain.setValueAtTime(0, now);
      gainNode.gain.linearRampToValueAtTime(0.5, now + 0.02);
      gainNode.gain.setValueAtTime(0.5, now + 0.1);
      gainNode.gain.linearRampToValueAtTime(0.01, now + 0.3);
      osc.start(now);
      osc.stop(now + 0.3);
    }
  } catch (error) {
    console.error("Gagal memutar audio:", error);
  }
};

// === NEW MODULE: DAILY SPEAKING CHALLENGE ===
function StudentSpeakingChallengeModule({ db, setDb, user, showToast, language = 'en' }) {
  // BANK DATA 90 HARI x 5 LEVEL (Triwulan)
  // Format Array: [Level 1 (TK), Level 2 (SD 1-2), Level 3 (SD 3-4), Level 4 (SD 5-6), Level 5 (SMP/SMA/Umum)]
  const SENTENCE_BANK = useMemo(() => [
    ["Apple", "I like apples", "I eat an apple every day", "Eating fruits like apples is very healthy", "An apple a day keeps the doctor away."],
    ["Red car", "It is a car", "The red car is very fast", "My father drives a red car to work", "Automobiles have revolutionized modern transportation drastically."],
    ["Cat", "The cat sleeps", "My cat likes to play outside", "I have a cute cat with white fur", "Felines are known for their remarkable agility and independence."],
    ["Book", "I read books", "Reading books makes me smart", "I borrowed a new book from the library", "Literature opens up entirely new worlds of human imagination."],
    ["Sun", "The sun shines", "The sun is yellow and hot", "The sun rises in the east every morning", "Solar energy is crucial for sustaining all life on Earth."],
    ["Water", "Drink some water", "I drink water every day", "Don't forget to drink eight glasses of water", "Hydration is essential for maintaining optimal bodily functions."],
    ["Happy", "I am happy", "Today I feel very happy", "Playing with my friends makes me so happy", "Happiness is often found in the simplest moments of life."],
    ["School", "I love school", "My school is big and clean", "I learn many new things at school today", "Education is the most powerful weapon to change the world."],
    ["Milk", "I drink milk", "Milk is good for my bones", "I always drink a glass of warm milk", "Dairy products provide calcium which is vital for skeletal health."],
    ["Blue sky", "The sky is blue", "Look at the beautiful blue sky", "There are no clouds in the blue sky", "The vast azure sky brings a sense of profound peace."],
    ["Dog", "The dog barks", "My dog runs in the park", "I walk my dog every afternoon after school", "Dogs have been faithful companions to humans for millennia."],
    ["Tree", "It is a tree", "The tree has many green leaves", "We should plant more trees to save earth", "Deforestation poses a significant threat to global ecosystems."],
    ["House", "This is my house", "My house has a small garden", "I live in a comfortable house with family", "A house is made of walls, but a home is made of love."],
    ["Bird", "The bird flies", "I see a bird in the tree", "The little bird is singing a beautiful song", "Avian creatures possess incredible navigational skills during migration."],
    ["Friend", "He is my friend", "I have many good friends here", "My friends and I always study together happily", "True friendship is a rare and invaluable treasure."],
    ["Food", "I eat food", "This food is very delicious", "My mother cooks delicious food for our dinner", "Culinary traditions reflect the rich cultural heritage of a nation."],
    ["Shoes", "My new shoes", "I wear my new black shoes", "I bought these shoes at the big mall", "Appropriate footwear is essential for athletic performance."],
    ["Rain", "It is raining", "The rain makes the grass wet", "I need an umbrella because it is raining", "Precipitation is a vital component of the global water cycle."],
    ["Pencil", "I have a pencil", "I write with a sharp pencil", "Please lend me your pencil for a moment", "The pen is often considered mightier than the sword."],
    ["Fish", "The fish swims", "The fish swims in the clear water", "I keep beautiful fish in a small aquarium", "Marine biodiversity is fascinating and crucial for the ocean."],
    ["Smile", "Please smile", "Your smile makes me feel happy", "A warm smile can easily brighten someone's day", "A genuine smile transcends all language barriers instantly."],
    ["Clock", "Look at the clock", "The clock shows it is time", "It is already ten o'clock in the morning", "Punctuality is a sign of respect for other people's time."],
    ["Teacher", "My teacher is kind", "I listen to my good teacher", "My English teacher explains the lesson very clearly", "Educators play a pivotal role in shaping future generations."],
    ["Chair", "Sit on the chair", "This wooden chair is very comfortable", "Please sit on the chair and wait here", "Ergonomic furniture prevents severe back pain over time."],
    ["Moon", "The moon is bright", "I see the moon at night", "The full moon shines brightly in the dark", "Lunar phases have influenced human calendars for centuries."],
    ["Run", "I can run", "I run fast in the field", "Running every morning keeps my body very healthy", "Cardiovascular exercises like running improve overall stamina."],
    ["Ball", "Kick the ball", "We play with a round ball", "He kicked the football straight into the goal", "Team sports foster cooperation and strategic thinking."],
    ["Flower", "A pretty flower", "The red flower smells very sweet", "I will give this beautiful flower to mom", "Botany is the scientific study of complex plant life."],
    ["Hand", "Wash your hands", "I wash my hands with soap", "You must wash your hands before eating lunch", "Personal hygiene is the first defense against infectious diseases."],
    ["Door", "Open the door", "Please open the front door now", "Remember to lock the door when you leave", "Opportunity rarely knocks twice on the same door."],
    ["Window", "Close the window", "The window is made of clear glass", "I look out the window to see the view", "Transparent architecture maximizes natural lighting in modern buildings."],
    ["Jump", "I like to jump", "The frog jumps very high", "Children love to jump up and down joyfully", "Plyometric exercises significantly enhance explosive muscle power."],
    ["Star", "Look at the star", "The stars are shining tonight", "There are millions of stars in the galaxy", "Astronomy seeks to understand the vastness of the universe."],
    ["Sing", "I can sing", "She sings a very beautiful song", "We will sing together at the school concert", "Vocal training requires immense discipline and breath control."],
    ["Sleep", "Time to sleep", "I sleep early every single night", "You should get enough sleep to stay healthy", "Adequate rest is fundamental for cognitive recovery."],
    ["Play", "Let's play together", "I play games with my brother", "Playing outdoor games is better than watching TV", "Recreational activities significantly reduce psychological stress."],
    ["Train", "It is a train", "The train moves on the track", "We travel to the big city by train", "Railway networks are the backbone of industrial logistics."],
    ["Boat", "The boat sails", "The wooden boat is on the sea", "Fishermen use boats to catch fish every day", "Maritime exploration opened up early global trade routes."],
    ["Hat", "Put on your hat", "I wear a hat in the sun", "This big hat protects my face from sunlight", "Appropriate headgear is necessary under extreme weather conditions."],
    ["Cold", "I feel cold", "The weather is very cold today", "Wear a jacket because it is cold outside", "Subzero temperatures demand adequate thermal insulation."],
    ["Hot", "It is very hot", "The hot soup burns my tongue", "Be careful because the cup of tea is hot", "Thermodynamics is the study of heat and energy transfer."],
    ["Clean", "Clean the room", "I help my mother clean the house", "Keeping our classroom clean is everyone's shared responsibility", "Sanitation protocols are strictly enforced in the facility."],
    ["Dirty", "My shoes are dirty", "Please wash your dirty hands now", "You should not wear those dirty clothes again", "Contaminated environments accelerate the spread of harmful pathogens."],
    ["Big", "A big elephant", "The elephant is a big animal", "They live in a very big and luxurious house", "The sheer magnitude of the project requires careful planning."],
    ["Small", "A small mouse", "The mouse is very small indeed", "I found a small kitten behind my house", "Microscopic organisms play a massive role in our ecosystem."],
    ["Fast", "Run very fast", "The cheetah can run very fast", "He finished his homework fast so he could play", "Rapid technological advancement characterizes the twenty-first century."],
    ["Slow", "A slow turtle", "The old turtle walks very slowly", "Please speak slowly so I can understand you", "A methodical approach guarantees higher accuracy in research."],
    ["Loud", "A loud noise", "The music is too loud here", "Please do not talk too loud in the library", "Excessive decibel levels can cause permanent auditory damage."],
    ["Quiet", "Be quiet please", "The library is a quiet place", "We must keep quiet while the baby sleeps", "Tranquility is often necessary for deep intellectual focus."],
    ["Good", "A good boy", "You did a very good job", "Eating vegetables is a good habit for us", "Moral integrity is the foundation of a civilized society."],
    ["Bad", "A bad dream", "I had a bad dream last night", "Eating too much candy is a bad habit", "Detrimental behaviors often lead to severe long-term consequences."],
    ["Easy", "It is easy", "This math test is very easy", "Learning English is easy if you practice daily", "Simplifying complex procedures improves overall workflow efficiency."],
    ["Hard", "It is hard", "This puzzle is too hard for me", "Working hard is the secret to achieving success", "Overcoming formidable challenges builds resilient character traits."],
    ["Early", "Wake up early", "I wake up early in the morning", "Arriving early gives you time to prepare yourself", "Proactive measures mitigate potential risks before they escalate."],
    ["Late", "Do not be late", "He was late for school today", "You should apologize for coming late to class", "Chronic tardiness often reflects a lack of professional discipline."],
    ["New", "My new toy", "I got a new bicycle today", "This new smartphone has a very good camera", "Innovation constantly drives the tech industry forward."],
    ["Old", "An old man", "My grandfather is an old man", "This old castle has been here for centuries", "Preserving historical artifacts maintains our connection to the past."],
    ["Right", "That is right", "Your answer to the question is right", "It is important to do the right thing", "Ethical reasoning distinguishes right from profound wrong."],
    ["Wrong", "That is wrong", "He realized his decision was completely wrong", "Do not be afraid to make a wrong guess", "Acknowledging miscalculations is the first step toward genuine growth."],
    ["High", "A high mountain", "The bird flies high in the sky", "They climbed the high mountain during the summer", "Elevated altitudes drastically affect atmospheric pressure levels."],
    ["Low", "A low table", "The water level is very low", "Please put the heavy box on the low table", "Subterranean environments possess unique ecological characteristics."],
    ["Far", "It is very far", "My school is far from my house", "We walked a far distance to reach the lake", "Remote geographic locations present significant logistical challenges."],
    ["Near", "It is near here", "The shop is near my home", "Fortunately, the hospital is near our new apartment", "Proximity to urban centers increases property valuation significantly."],
    ["Full", "The glass is full", "My stomach is full after eating", "The bus was full so we waited for another", "Maximum capacity has been reached within the main auditorium."],
    ["Empty", "An empty box", "The water bottle is completely empty", "The empty room echoed when I spoke loudly", "A void environment is necessary for certain physics experiments."],
    ["Light", "A light feather", "The small bag is very light", "Turn on the light so we can read", "Photonic technology is advancing at an unprecedented rate."],
    ["Heavy", "A heavy rock", "This box is too heavy for me", "I need your help to lift this heavy desk", "Industrial machinery is designed to handle exceptionally massive loads."],
    ["Soft", "A soft pillow", "My bed has a soft pillow", "She spoke to me in a very soft voice", "A diplomatic approach often yields more constructive negotiations."],
    ["Hard", "A hard stone", "The rock is very hard indeed", "It rained hard all night without stopping once", "Enduring severe hardship is a testament to human resilience."],
    ["Rich", "A rich man", "The rich man bought a big car", "A rich vocabulary helps you speak English fluently", "Economic prosperity often correlates with advanced infrastructural development."],
    ["Poor", "A poor boy", "We should help the poor people", "Poor time management will ruin your daily schedule", "Systemic poverty requires multifaceted socio-economic interventions."],
    ["Beautiful", "A beautiful girl", "The sunset is very beautiful today", "She wore a beautiful dress to the grand party", "Aesthetic appreciation varies greatly across different cultural paradigms."],
    ["Ugly", "An ugly monster", "The monster in the story looks ugly", "Using bad words is an ugly habit to have", "Superficial judgments often overlook inherent intrinsic value."],
    ["Strong", "A strong man", "He is strong enough to lift it", "You need a strong password to protect your account", "Robust cybersecurity protocols are essential for data protection."],
    ["Weak", "I feel weak", "The sick boy feels very weak", "A weak foundation will cause the building to fall", "Structural vulnerabilities compromise the integrity of the entire system."],
    ["Brave", "A brave hero", "The brave knight saved the village", "It takes a brave person to admit their mistakes", "Courage is not the absence of fear, but mastering it."],
    ["Afraid", "Do not be afraid", "I am afraid of the dark", "There is no reason to be afraid of trying", "Irrational phobias can severely hinder personal psychological development."],
    ["Smart", "A smart student", "He is a very smart boy", "Smart students always ask questions when they are confused", "Intellectual curiosity drives the pursuit of academic excellence."],
    ["Stupid", "Do not be stupid", "It was a stupid mistake indeed", "Nobody is stupid, everyone just learns at different speeds", "Ignorance is easily cured by a genuine willingness to learn."],
    ["Funny", "A funny clown", "The clown is very funny today", "He told a funny joke that made everyone laugh", "Humor serves as an excellent psychological coping mechanism."],
    ["Sad", "I am sad", "The sad movie made me cry", "She looked sad because she lost her favorite pen", "Melancholy is a natural emotional response to profound loss."],
    ["Angry", "He is angry", "My mother was angry at me", "Please do not be angry about this small mistake", "Uncontrolled hostility often exacerbates already delicate diplomatic situations."],
    ["Tired", "I am tired", "I feel tired after running fast", "You look tired, you should get some rest now", "Chronic fatigue syndrome drastically diminishes overall operational productivity."],
    ["Hungry", "I am hungry", "I am hungry, let's eat now", "The hungry dog barked loudly begging for some food", "Malnutrition remains a critical issue in developing global regions."],
    ["Thirsty", "I am thirsty", "I need water, I am thirsty", "Drinking a cold beverage is perfect when you're thirsty", "Dehydration can lead to severe physiological complications if ignored."],
    ["Busy", "I am busy", "My father is very busy today", "The busy street was filled with cars and motorcycles", "A congested schedule necessitates rigorous prioritization and delegation."],
    ["Lazy", "Do not be lazy", "The lazy cat sleeps all day", "Being lazy will stop you from reaching your dreams", "Procrastination is the thief of time and potential success."],
    ["Careful", "Be careful please", "Be careful when crossing the road", "A careful driver always stops at the red light", "Meticulous attention to detail prevents catastrophic systemic failures."],
    ["Careless", "A careless mistake", "He lost his phone because he was careless", "Careless actions can hurt you and the people around", "Negligence in protocol execution often triggers severe legal liabilities."],
    // --- BEGIN 275 NEW DAILY CHALLENGES (365 DAYS GOD TIER) ---
    ["Yellow", "A yellow sun", "The sun is yellow", "I have a yellow pencil in my bag", "Yellow wavelengths are easily perceived by the human eye."],
    ["Green", "Green grass", "The grass is green", "I like to sit on the green grass", "Chlorophyll gives plants their distinct green pigmentation."],
    ["Pink", "A pink dress", "She wears a pink dress", "Pink is her favorite color for clothes", "Cultural associations often link pink to subtle affection."],
    ["Purple", "A purple flower", "This flower is purple", "The queen wears a long purple cloak", "Historically, purple dye was a symbol of royal exclusivity."],
    ["Orange", "An orange fruit", "The orange is sweet", "I drink a glass of fresh orange juice", "Citrus fruits provide essential vitamins for the immune system."],
    ["Brown", "A brown bear", "The big bear is brown", "The wooden table is painted dark brown", "Earthy tones create a warm and inviting architectural ambiance."],
    ["Black", "A black cat", "The cat is black", "He drives a shiny black car today", "Black absorbs all visible wavelengths of the light spectrum."],
    ["White", "White snow", "The snow is white", "Look at the beautiful white clouds above", "A completely white canvas represents infinite artistic possibilities."],
    ["Gray", "A gray stone", "The rock is gray", "The sky turns gray before the rain", "Monochrome palettes offer a minimalist aesthetic in modern design."],
    ["Silver", "A silver coin", "This coin is silver", "My mother wears a beautiful silver ring", "Silver possesses the highest electrical conductivity of any element."],
    ["One", "Number one", "I have one apple", "He is the number one student here", "Singularity represents the most fundamental mathematical concept."],
    ["Two", "Two eyes", "I have two eyes", "We need two tickets for the movie", "Bilateral symmetry is extremely common in the animal kingdom."],
    ["Three", "Three birds", "I see three birds", "There are three tall trees in the park", "A triangle is the simplest two-dimensional geometric polygon."],
    ["Four", "Four legs", "A dog has four legs", "A car has four wheels to move", "The four cardinal directions are essential for terrestrial navigation."],
    ["Five", "Five fingers", "I have five fingers", "Give me a high five for good luck", "Pentagonal structures are frequently observed in molecular chemistry."],
    ["Ten", "Ten toes", "I have ten toes", "She counted from one to ten slowly", "The decimal system is the standard for numerical representation."],
    ["Hundred", "One hundred", "A hundred days", "He has a hundred books in his room", "Centuries dictate the way historians categorize human epochs."],
    ["Thousand", "A thousand stars", "Thousands of people", "I saw a thousand stars in the sky", "Millennia encompass vast stretches of evolutionary geological time."],
    ["First", "The first day", "I am the first", "She won first place in the big race", "Initial impressions significantly influence subsequent interpersonal relationships."],
    ["Last", "The last page", "He is the last", "This is the last piece of the cake", "Final conclusions must be drawn from empirical scientific data."],
    ["Tiger", "A fierce tiger", "The tiger is orange", "The strong tiger runs in the wild jungle", "Tigers are apex predators crucial for maintaining ecological balance."],
    ["Lion", "The brave lion", "The lion is sleeping", "The lion is known as the king of beasts", "Prides of lions exhibit complex hierarchical social structures."],
    ["Bear", "A big bear", "The bear loves honey", "The brown bear catches fish in the river", "Hibernation allows bears to survive extreme winter conditions."],
    ["Monkey", "A funny monkey", "The monkey jumps high", "The clever monkey swings from tree to tree", "Primates display remarkable problem-solving and tool-using capabilities."],
    ["Zebra", "A striped zebra", "The zebra is fast", "Zebras have black and white stripes on them", "Unique stripe patterns serve as camouflage against predatory vision."],
    ["Giraffe", "A tall giraffe", "The giraffe eats leaves", "A giraffe has a very long neck", "The elongated cervical vertebrae of giraffes are evolutionary marvels."],
    ["Penguin", "A cute penguin", "The penguin likes cold", "Penguins walk funny but swim very fast", "Avian adaptations allow penguins to thrive in subzero environments."],
    ["Shark", "A big shark", "The shark has teeth", "The great white shark swims in the ocean", "Cartilaginous skeletons provide sharks with incredible hydrodynamic efficiency."],
    ["Whale", "A huge whale", "The whale is big", "The blue whale is the largest animal ever", "Cetaceans communicate using complex low-frequency acoustic vocalizations."],
    ["Dolphin", "A smart dolphin", "The dolphin jumps up", "Dolphins love to play in the blue sea", "Echolocation is a sophisticated biological sonar utilized by dolphins."],
    ["Snake", "A long snake", "The snake is green", "The poisonous snake slithers in the tall grass", "Ophidiophobia is the psychological fear associated with limbless reptiles."],
    ["Frog", "A green frog", "The frog jumps far", "I heard a frog croaking in the pond", "Amphibian life cycles involve fascinating metamorphic developmental stages."],
    ["Rabbit", "A soft rabbit", "The rabbit eats carrots", "My white rabbit hops around the garden fast", "Lagomorphs possess continuously growing incisors requiring constant chewing."],
    ["Spider", "A small spider", "The spider makes webs", "The spider caught a fly in its web", "Arachnids synthesize high-tensile silk for intricate structural webs."],
    ["Horse", "A fast horse", "The horse runs fast", "He rides a brown horse in the farm", "Equestrian sports demand exceptional synchronization between rider and mount."],
    ["Cow", "A fat cow", "The cow gives milk", "The black and white cow eats green grass", "Bovine livestock are fundamental to the global agricultural economy."],
    ["Pig", "A pink pig", "The pig is muddy", "The cute pig sleeps in the warm barn", "Porcine intelligence is often significantly underestimated by the public."],
    ["Sheep", "A fluffy sheep", "The sheep is white", "The farmer shears the thick wool from sheep", "Ovine herds are historically pivotal to the textile industry."],
    ["Goat", "A horned goat", "The goat climbs up", "The goat eats leaves from the small bush", "Caprine adaptability allows them to inhabit treacherous mountainous terrains."],
    ["Chicken", "A loud chicken", "The chicken lays eggs", "My grandmother feeds the chicken in the yard", "Poultry farming provides a substantial portion of global dietary protein."],
    ["Bread", "Eat the bread", "I like sweet bread", "I ate some bread with butter for breakfast", "Fermentation processes using yeast are essential in baking bread."],
    ["Cheese", "Yellow cheese", "The mouse likes cheese", "I put a slice of cheese on my burger", "Dairy coagulation techniques produce diverse global cheese varieties."],
    ["Butter", "Yellow butter", "Melt the soft butter", "Spread some butter on the hot toast please", "Lipid separation from cream creates this rich culinary ingredient."],
    ["Cake", "A sweet cake", "The cake is big", "My mother baked a chocolate cake for me", "Confectionery arts require precise measurements of basic baking ingredients."],
    ["Pizza", "A round pizza", "I love cheese pizza", "We ordered a large pizza for the party", "Gastronomic globalization has made pizza an internationally recognized staple."],
    ["Burger", "A big burger", "I eat a burger", "The beef burger has tomato and lettuce inside", "Fast food franchises have standardized burger production worldwide."],
    ["Rice", "White rice", "I eat warm rice", "We eat fried rice for dinner every day", "Cultivation of paddy fields is essential in Asian agriculture."],
    ["Noodle", "Long noodles", "I slurp the noodles", "This bowl of hot noodles is very spicy", "Extrusion techniques are historically utilized to manufacture various noodles."],
    ["Salad", "A green salad", "I mix the salad", "Eating a fresh salad is good for health", "Nutritional diets often incorporate raw vegetables to maximize vitamins."],
    ["Soup", "Hot soup", "The soup is warm", "I drink chicken soup when I feel sick", "Simmering broths extract complex flavors from bones and vegetables."],
    ["Meat", "Cook the meat", "The meat is red", "My father grills the meat outside the house", "Protein synthesis in the human body requires dietary amino acids."],
    ["Egg", "An oval egg", "I boil an egg", "I eat a fried egg in the morning", "Avian ovulation provides a highly versatile and nutritious food source."],
    ["Mango", "A sweet mango", "The mango is yellow", "I love eating ripe mangoes in the summer", "Tropical climates are optimal for the cultivation of mango orchards."],
    ["Banana", "A yellow banana", "Monkeys like yellow bananas", "I peeled a fresh banana for my snack", "Potassium-rich fruits like bananas are excellent for muscular recovery."],
    ["Lemon", "A sour lemon", "The lemon is sour", "I squeeze a fresh lemon into my tea", "Ascorbic acid gives citrus fruits their characteristic tart flavor."],
    ["Grape", "Purple grapes", "I wash the grapes", "These purple grapes are very sweet and juicy", "Viticulture is the specialized agricultural study of cultivating grapevines."],
    ["Carrot", "An orange carrot", "Rabbits eat orange carrots", "Eating carrots is very good for your eyes", "Beta-carotene is a crucial antioxidant found abundantly in carrots."],
    ["Potato", "A brown potato", "I mash the potato", "French fries are made from freshly cut potatoes", "Tubers serve as critical carbohydrate reserves for many global populations."],
    ["Tomato", "A red tomato", "The tomato is red", "Slice the red tomato for the fresh salad", "Botanically classified as fruits, tomatoes are culinary vegetables worldwide."],
    ["Onion", "A round onion", "The onion makes me cry", "Chop the white onion for the hot soup", "Sulfenic acid release causes ocular irritation when chopping onions."],
    ["Mountain", "A high mountain", "The mountain is green", "We will climb the high mountain next week", "Tectonic plate collisions are responsible for massive geological uplifts."],
    ["Hill", "A green hill", "The hill is steep", "The children rolled down the grassy green hill", "Topographical elevations like hills offer strategic geographical vantage points."],
    ["River", "A long river", "The river flows fast", "We caught a big fish in the river", "Fluvial ecosystems support an incredible diversity of aquatic species."],
    ["Lake", "A calm lake", "The lake is blue", "We rented a small boat on the lake", "Lentic environments provide crucial freshwater reservoirs for surrounding biomes."],
    ["Sea", "The deep sea", "The sea is salty", "I love swimming in the clear blue sea", "Marine currents significantly influence global meteorological and climatic patterns."],
    ["Ocean", "The vast ocean", "The ocean is huge", "Whales and dolphins live in the deep ocean", "Oceanic expanses cover approximately seventy-one percent of planetary surfaces."],
    ["Beach", "A sandy beach", "We play at the beach", "I built a sandcastle on the sunny beach", "Coastal erosion is a continuous environmental challenge for shoreline communities."],
    ["Sand", "Yellow sand", "The sand is hot", "I dug a deep hole in the sand", "Granular silica forms the primary composition of most terrestrial deserts."],
    ["Island", "A small island", "The island is beautiful", "We took a boat to the remote island", "Isolated landmasses frequently foster unique evolutionary biological developments."],
    ["Forest", "A dark forest", "The forest has trees", "Many wild animals live in the thick forest", "Arboreal ecosystems act as vital carbon sinks for the planet."],
    ["Jungle", "A wild jungle", "The jungle is green", "Explorers discovered a hidden temple in the jungle", "Tropical rainforests boast unparalleled levels of terrestrial biodiversity globally."],
    ["Desert", "A dry desert", "The desert is hot", "Camels can survive in the hot dry desert", "Arid landscapes receive minimal annual precipitation, challenging biological survival."],
    ["Cave", "A dark cave", "The cave is scary", "Bats sleep hanging upside down in the cave", "Speleology involves the scientific exploration and study of subterranean caverns."],
    ["Rock", "A hard rock", "The rock is heavy", "He threw a small rock into the pond", "Geological formations consist of various aggregated mineral compositions over time."],
    ["Cloud", "A white cloud", "The cloud is fluffy", "Dark clouds usually mean it will rain soon", "Meteorological condensation forms visible masses of suspended atmospheric water droplets."],
    ["Wind", "A strong wind", "The wind blows hard", "The cold wind made the tree branches sway", "Atmospheric pressure differentials generate kinetic energy known as wind."],
    ["Storm", "A bad storm", "The storm is scary", "We stayed indoors because of the terrible storm", "Severe meteorological disturbances can cause catastrophic infrastructural and environmental damage."],
    ["Snow", "Cold white snow", "The snow is falling", "We built a big snowman in the yard", "Crystalline water ice forms intricate hexagonal structures during atmospheric freezing."],
    ["Ice", "Cold hard ice", "The ice is slippery", "I put some ice cubes in my drink", "Solidified states of water undergo unique volumetric expansion upon freezing."],
    ["Rainbow", "A colorful rainbow", "Look at the rainbow", "A beautiful rainbow appeared right after the rain", "Optical refraction and dispersion create meteorological phenomena spanning visible spectrums."],
    ["Table", "A wooden table", "Put it on the table", "We eat our dinner at the dining table", "Furniture design often prioritizes both structural stability and aesthetic appeal."],
    ["Desk", "A school desk", "I sit at my desk", "She left her notebook on the wooden desk", "Ergonomic workspaces significantly enhance professional and academic productivity overall."],
    ["Sofa", "A soft sofa", "Sit on the sofa", "We sat on the comfortable sofa watching TV", "Upholstered seating arrangements provide maximum comfort within residential living areas."],
    ["Bed", "A soft bed", "I sleep in my bed", "Make your bed after you wake up morning", "Mattress support systems are crucial for maintaining proper spinal alignment."],
    ["Lamp", "A bright lamp", "Turn on the lamp", "I need a reading lamp for my desk", "Artificial illumination extends operational hours beyond natural solar availability."],
    ["Mirror", "Look in the mirror", "The mirror is clean", "She brushed her hair in front of mirror", "Reflective surfaces utilize polished metallic coatings to bounce light precisely."],
    ["Picture", "A nice picture", "Look at this picture", "I hung a beautiful family picture on wall", "Photographic memories capture transient moments for permanent historical preservation."],
    ["TV", "Watch the TV", "The TV is loud", "We watched an exciting movie on the TV", "Broadcasting networks transmit audiovisual signals to millions of global households."],
    ["Radio", "Listen to radio", "The radio plays music", "My grandfather listens to news on the radio", "Electromagnetic frequency modulation allows for wireless transmission of acoustic data."],
    ["Computer", "A fast computer", "I use the computer", "He plays video games on his personal computer", "Microprocessors execute complex algorithms to facilitate advanced computational tasks."],
    ["Laptop", "A new laptop", "Close the laptop", "I bring my thin laptop to the cafe", "Portable computing devices revolutionized modern remote professional working environments."],
    ["Phone", "A smart phone", "My phone is ringing", "She sent a text message using her phone", "Telecommunication advancements have fundamentally restructured global interpersonal connectivity."],
    ["Plate", "A clean plate", "Put food on the plate", "Wash the dirty plate after you finish eating", "Ceramic dinnerware is manufactured through high-temperature kiln firing processes."],
    ["Bowl", "A deep bowl", "A bowl of soup", "I ate a hot bowl of chicken soup", "Concave vessels are optimally designed for containing liquid culinary preparations."],
    ["Cup", "A red cup", "A cup of tea", "Please pour some hot water into my cup", "Thermal insulation in mugs prevents heat transfer from hot beverages."],
    ["Glass", "A clear glass", "A glass of water", "Be careful not to break the fragile glass", "Silicate compounds undergo rapid cooling to form transparent amorphous solids."],
    ["Spoon", "A silver spoon", "Eat with a spoon", "Use a spoon to eat the hot soup", "Cutlery design balances functional utility with ergonomic human-centric handling."],
    ["Fork", "A sharp fork", "Use the fork", "He eats the delicious pasta with a fork", "Tined utensils facilitate the piercing and stabilization of solid foods."],
    ["Knife", "A sharp knife", "Cut with the knife", "Use the sharp knife to slice the meat", "Metallurgical hardening ensures the durability of culinary cutting edge implements."],
    ["Pot", "A hot pot", "Cook in the pot", "My mother boiled water in the large pot", "Thermal conductivity in metallic cookware ensures even heat distribution during cooking."],
    ["Bank", "A big bank", "Go to the bank", "I need to withdraw some money from bank", "Financial institutions regulate monetary circulation and facilitate macroeconomic stability."],
    ["Post Office", "The post office", "Mail the letter", "I sent a package at the post office", "Logistical networks efficiently manage the global distribution of physical parcels."],
    ["Hospital", "A clean hospital", "Go to the hospital", "The ambulance rushed the patient to the hospital", "Medical facilities provide critical infrastructural support for complex surgical procedures."],
    ["Clinic", "A small clinic", "Visit the clinic", "She went to the dental clinic for a checkup", "Outpatient healthcare centers specialize in localized preventative and diagnostic medicine."],
    ["Police Station", "The police station", "Call the police", "The officer took the thief to police station", "Law enforcement agencies maintain societal order and investigate criminal activities systematically."],
    ["Fire Station", "The fire station", "Call the firemen", "The red fire truck left the fire station", "Emergency response units coordinate rapid deployment to mitigate catastrophic conflagrations."],
    ["Bakery", "A sweet bakery", "Buy some bread", "I bought fresh warm bread at the bakery", "Artisanal baking requires precise mastery of complex dough fermentation processes."],
    ["Supermarket", "A big supermarket", "Buy some food", "We bought groceries for the week at supermarket", "Retail conglomerates optimize supply chains to ensure massive inventory availability."],
    ["Market", "A busy market", "Go to the market", "The traditional market sells very fresh green vegetables", "Decentralized commerce hubs foster direct economic transactions between local producers."],
    ["Mall", "A huge mall", "Shop at the mall", "We spent the weekend shopping at the mall", "Commercial complexes integrate retail, dining, and entertainment into unified structures."],
    ["Cinema", "A dark cinema", "Watch a movie", "We bought popcorn before entering the dark cinema", "Theatrical venues utilize advanced acoustic engineering for immersive audiovisual experiences."],
    ["Theater", "A grand theater", "Watch the play", "The actors performed brilliantly on the theater stage", "Dramaturgical productions require meticulous coordination of lighting, sound, and choreography."],
    ["Museum", "An old museum", "Look at the art", "We saw ancient dinosaur bones at the museum", "Curatorial institutions preserve and exhibit artifacts of profound historical significance."],
    ["Zoo", "A fun zoo", "See the animals", "The children loved seeing the elephants at zoo", "Zoological parks contribute significantly to global wildlife conservation and research."],
    ["Park", "A green park", "Play in the park", "We had a lovely picnic at the park", "Urban green spaces provide essential recreational areas mitigating metropolitan density."],
    ["Airport", "A busy airport", "Fly from the airport", "We waited for our flight at the airport", "Aviation hubs manage the complex logistical choreography of international aerial transit."],
    ["Station", "A train station", "Wait at the station", "The train arrived late at the central station", "Terminal infrastructure facilitates the efficient boarding of mass transit commuter systems."],
    ["Port", "A big port", "Ships at the port", "Massive cargo ships dock at the busy port", "Maritime harbors are the primary nodes for international freight logistics operations."],
    ["Hotel", "A luxury hotel", "Sleep at the hotel", "We booked a room at a nice hotel", "Hospitality enterprises prioritize exceptional service standards to ensure guest satisfaction."],
    ["Restaurant", "A nice restaurant", "Eat at the restaurant", "They celebrated their anniversary at a fancy restaurant", "Culinary establishments synthesize gastronomic expertise with sophisticated ambient dining experiences."],
    ["Doctor", "A kind doctor", "The doctor helps", "The doctor gave me medicine for my fever", "Medical practitioners dedicate their lives to diagnosing and treating complex pathologies."],
    ["Nurse", "A gentle nurse", "The nurse cares", "The nurse checked my blood pressure very carefully", "Clinical caregivers provide indispensable continuous monitoring and empathetic patient support."],
    ["Dentist", "A smart dentist", "Check your teeth", "The dentist cleaned my teeth and fixed cavities", "Odontological specialists focus exclusively on the maintenance of oral hygiene."],
    ["Policeman", "A brave policeman", "The policeman runs", "The policeman helped the lost child find home", "Law enforcement officers undergo rigorous tactical training to ensure public safety."],
    ["Firefighter", "A strong firefighter", "Put out the fire", "The brave firefighter rescued the cat from flames", "Emergency responders courageously confront hazardous environments to preserve civilian lives."],
    ["Postman", "A busy postman", "Deliver the mail", "The postman delivered a heavy package today morning", "Courier personnel are integral to the reliable functioning of communication networks."],
    ["Driver", "A careful driver", "Drive the car", "The bus driver drove us safely to school", "Professional chauffeurs must strictly adhere to comprehensive vehicular safety regulations."],
    ["Pilot", "A smart pilot", "Fly the plane", "The pilot landed the airplane smoothly on runway", "Aeronautical commanders navigate sophisticated avionic systems during complex atmospheric flights."],
    ["Farmer", "A hard farmer", "Plant the seeds", "The farmer grows fresh vegetables in the field", "Agricultural cultivators implement sustainable techniques to maximize seasonal crop yields."],
    ["Cook", "A good cook", "Cook the food", "The cook prepared a delicious meal for us", "Culinary professionals consistently execute complex recipes under intense temporal pressure."],
    ["Chef", "A master chef", "The chef cooks", "The head chef created a brilliant new menu", "Executive gastronomists innovate flavor profiles while managing rigorous kitchen operations."],
    ["Waiter", "A polite waiter", "Serve the food", "The waiter brought our drinks to the table", "Hospitality staff ensure seamless communicative coordination between patrons and kitchens."],
    ["Baker", "A busy baker", "Bake the bread", "The baker woke up early to make pastries", "Artisanal confectioners master the delicate biochemical reactions required for baking."],
    ["Student", "A smart student", "Study hard everyday", "The diligent student always completes his homework perfectly", "Academic scholars must cultivate profound analytical and critical thinking competencies."],
    ["Singer", "A loud singer", "Sing a song", "The famous singer performed beautifully on the stage", "Vocal artists undergo strenuous physiological conditioning to perfect acoustic resonance."],
    ["Actor", "A dramatic actor", "Act in a play", "The lead actor won an award for acting", "Theatrical performers meticulously embody complex psychological profiles for dramatic authenticity."],
    ["Artist", "A creative artist", "Paint a picture", "The talented artist painted a gorgeous landscape canvas", "Visual creatives manipulate spatial dimensions to evoke profound emotional aesthetic responses."],
    ["Writer", "A quiet writer", "Write a book", "The author is writing a new mystery novel", "Literary composers weave intricate narrative structures to explore profound human conditions."],
    ["Scientist", "A smart scientist", "Do an experiment", "The brilliant scientist discovered a cure for the disease", "Empirical researchers formulate rigorous hypotheses to decode complex universal phenomena."],
    ["Engineer", "A clever engineer", "Build a bridge", "The structural engineer designed a very safe bridge", "Technological innovators apply advanced mathematical paradigms to construct robust infrastructural solutions."],
    ["Bus", "A yellow bus", "Ride the bus", "I take the school bus every single morning", "Municipal transit systems alleviate severe metropolitan traffic congestion significantly overall."],
    ["Taxi", "A fast taxi", "Call a taxi", "We took a yellow taxi to the airport", "Chauffeured transport offers personalized locational flexibility within complex urban environments."],
    ["Bicycle", "A red bicycle", "Ride your bicycle", "I ride my bicycle to the park everyday", "Non-motorized pedaled transport provides excellent cardiovascular benefits and zero emissions."],
    ["Motorcycle", "A fast motorcycle", "Ride the motorcycle", "He wears a helmet when riding his motorcycle", "Two-wheeled vehicular transit requires exceptional kinetic balance and spatial awareness."],
    ["Airplane", "A big airplane", "Fly in airplane", "The massive airplane flew high above the clouds", "Aerodynamic propulsion allows massive commercial jets to traverse intercontinental distances."],
    ["Helicopter", "A loud helicopter", "Fly the helicopter", "The rescue helicopter landed safely on the roof", "Rotary-wing aircraft possess unique vertical takeoff and precise hovering capabilities."],
    ["Truck", "A heavy truck", "Drive the truck", "The delivery truck carried heavy boxes of food", "Heavy-duty freight logistics heavily depend on robust diesel-powered vehicular transport."],
    ["Van", "A white van", "Drive the van", "Our family rented a large van for vacation", "Versatile transport vehicles provide optimal spatial capacity for diverse logistical requirements."],
    ["Ambulance", "A fast ambulance", "Call the ambulance", "The ambulance sirens blared loudly through the streets", "Paramedic transit vehicles are equipped with sophisticated mobile life-support technologies."],
    ["Ship", "A huge ship", "Sail the ship", "The massive cargo ship sailed across the ocean", "Naval engineering ensures the structural integrity of colossal oceanic freight vessels."],
    ["Head", "My round head", "Nod your head", "I wear a warm hat on my head", "The cranial structure houses and protects the delicate neurological cerebral cortex."],
    ["Face", "A happy face", "Wash your face", "She has a beautiful smile on her face", "Facial micro-expressions subtly communicate a vast spectrum of complex human emotions."],
    ["Eye", "A blue eye", "Close your eyes", "I use my two eyes to see clearly", "Ocular photoreceptors translate environmental light spectrums into comprehensive visual data."],
    ["Ear", "A small ear", "Listen with ears", "I hear the loud music with my ears", "Auditory canals channel acoustic vibrations toward the intricate inner tympanic membrane."],
    ["Nose", "A pointy nose", "Smell with nose", "I use my nose to smell the flowers", "Olfactory sensors detect microscopic airborne chemical compounds translating into distinct scents."],
    ["Mouth", "A wide mouth", "Open your mouth", "I use my mouth to eat and speak", "Oral cavities serve as the primary gateway for both nutritional and phonetic functions."],
    ["Tooth", "A white tooth", "Brush your teeth", "I visit the dentist to check my teeth", "Enamel-coated dental structures are essential for the mechanical breakdown of sustenance."],
    ["Hair", "Long black hair", "Comb your hair", "She tied her long hair with a ribbon", "Follicular keratin filaments provide both thermal insulation and distinct aesthetic characteristics."],
    ["Neck", "A long neck", "Turn your neck", "She wore a shiny diamond necklace around neck", "Cervical vertebrae provide crucial structural support and rotational flexibility for craniums."],
    ["Shoulder", "A strong shoulder", "Touch your shoulder", "He carried the heavy backpack on his shoulders", "The complex deltoid musculature facilitates a remarkable multidirectional range of arm motion."],
    ["Arm", "A long arm", "Raise your arm", "I use my arms to hug my mother", "Appendicular skeletal extensions are biologically optimized for sophisticated manipulative environmental interactions."],
    ["Leg", "A strong leg", "Kick with leg", "I use my strong legs to run fast", "Lower extremities are anatomically engineered to sustain bipedal locomotion and equilibrium."],
    ["Foot", "A left foot", "Stomp your feet", "I wear comfortable soft shoes on my feet", "Pedal structures distribute gravitational weight enabling complex ambulatory propulsion mechanics."],
    ["Finger", "A small finger", "Point your finger", "I wear a beautiful gold ring on finger", "Phalangeal digits possess extreme tactile sensitivity enabling highly intricate manipulative dexterity."],
    ["Toe", "A little toe", "Wiggle your toes", "I stubbed my little toe on the table", "Terminal pedal appendages are critical for maintaining advanced kinetic upright balance."],
    ["Shirt", "A blue shirt", "Wear a shirt", "He ironed his white shirt for the meeting", "Textile garments provide essential environmental protection and reflect diverse cultural aesthetics."],
    ["T-shirt", "A red T-shirt", "Wear a T-shirt", "I bought a comfortable cotton T-shirt yesterday afternoon", "Casual apparel manufacturing frequently utilizes highly breathable and flexible synthetic polymers."],
    ["Pants", "Long blue pants", "Wear your pants", "These denim pants are too tight for me", "Bifurcated lower garments are specifically tailored to maximize bilateral ambulatory freedom."],
    ["Skirt", "A short skirt", "Wear a skirt", "She twirled around in her beautiful floral skirt", "Sartorial traditions often utilize unbifurcated fabrics to prioritize stylistic fluid movement."],
    ["Dress", "A pretty dress", "Wear a dress", "The bride wore a gorgeous white wedding dress", "Haute couture designs meticulously integrate elegant textiles to flatter human anatomical silhouettes."],
    ["Jacket", "A warm jacket", "Wear a jacket", "Put on your thick jacket because it's freezing", "Thermally insulated outerwear mitigates the physiological impact of severe subzero atmospheric conditions."],
    ["Coat", "A heavy coat", "Wear a coat", "He hung his wet winter coat to dry", "Overgarment tailoring incorporates specialized water-resistant textiles to combat inclement meteorological phenomena."],
    ["Sweater", "A soft sweater", "Wear a sweater", "My grandmother knitted a warm woolen sweater for me", "Interlocking knitted yarns provide exceptional thermal retention during severe seasonal temperature drops."],
    ["Socks", "Warm white socks", "Wear your socks", "I need to buy a new pair of socks", "Hosiery prevents abrasive dermal friction between pedal extremities and rigid footwear interiors."],
    ["Boots", "Heavy black boots", "Wear your boots", "He wore heavy boots to hike the mountain", "Reinforced tactical footwear ensures absolute stabilization traversing treacherous and unforgiving geological topographies."],
    ["Gloves", "Warm soft gloves", "Wear your gloves", "She wore thick leather gloves to ride motorcycle", "Manual protective gear prevents acute frostbite and severe epidermal laceration risks."],
    ["Scarf", "A long scarf", "Wear a scarf", "She wrapped a red scarf around her neck", "Cervical textile accessories efficiently mitigate rapid thermal dissipation from exposed jugular veins."],
    ["Tie", "A neat tie", "Wear a tie", "He adjusted his silk tie before the interview", "Sartorial accessories function primarily to signify strict adherence to formal corporate protocols."],
    ["Belt", "A leather belt", "Wear a belt", "He tightened his brown leather belt one notch", "Circumferential structural accessories maintain the precise positional integrity of lower torso garments."],
    ["Cap", "A cool cap", "Wear a cap", "He wore a baseball cap to block sun", "Visored headgear is ergonomically designed to deflect excessive ocular ultraviolet radiation exposure."],
    ["Walk", "I can walk", "Walk to school", "We took a long walk in the park", "Bipedal locomotion represents a critical evolutionary milestone for early terrestrial hominids."],
    ["Talk", "We can talk", "Talk to me", "The teacher told the students not to talk", "Verbal communication facilitates the complex interpersonal transmission of abstract conceptual ideologies."],
    ["Speak", "I can speak", "Speak in English", "She can speak three different languages very fluently", "Phonetic articulation requires precise synchronous coordination of complex neurological and muscular systems."],
    ["Listen", "I can listen", "Listen to music", "You must listen carefully to the given instructions", "Auditory processing involves deciphering complex acoustic frequencies into coherent semantic comprehension."],
    ["Hear", "I can hear", "Hear the sound", "I can hear the birds singing outside clearly", "Acoustic perception alerts biological organisms to immediate shifts within their surrounding environment."],
    ["Look", "Look at me", "Look at that", "Look both ways before you cross the street", "Ocular focus dictates the primary trajectory of conscious human environmental observation."],
    ["See", "I can see", "See the bird", "I cannot see anything without my reading glasses", "Visual cortex interpretation synthesizes photonic inputs into comprehensive cognitive spatial maps."],
    ["Watch", "Watch the TV", "Watch a movie", "We will watch a thrilling action movie tonight", "Sustained optical observation allows for the detailed analysis of sequential kinetic events."],
    ["Touch", "Touch the wall", "Do not touch", "Please do not touch the fragile museum artifacts", "Tactile stimulation triggers immediate neurological responses regarding surface texture and temperature."],
    ["Smell", "Smell the flower", "It smells good", "The kitchen smells amazing when mother is cooking", "Olfactory receptors are fundamentally linked to deeply ingrained neurological memory associations."],
    ["Taste", "Taste the food", "It tastes sweet", "I want to taste the delicious chocolate cake", "Gustatory perception distinguishes chemical compositions to ensure the safety of nutritional ingestion."],
    ["Bite", "Bite the apple", "Do not bite", "The hungry dog tried to bite the bone", "Mandibular exertion is the primary mechanical action required for initial digestive processing."],
    ["Sip", "Sip the water", "Sip the tea", "She took a small sip of hot coffee", "Controlled oral suction allows for the cautious ingestion of potentially hazardous thermal liquids."],
    ["Cook", "Cook the food", "My mom cooks", "He is learning to cook healthy meals everyday", "Thermal application induces complex biochemical alterations enhancing both flavor and nutritional safety."],
    ["Bake", "Bake a cake", "Bake the bread", "We will bake delicious cookies for the party", "Oven convection techniques rely on precise temperature regulation for optimal dough expansion."],
    ["Write", "Write a word", "Write your name", "I need a pen to write this letter", "Linguistic documentation requires the systematic physical transcription of abstract phonetic symbols."],
    ["Draw", "Draw a cat", "Draw a picture", "The clever child likes to draw colorful animals", "Illustrative drafting translates abstract cognitive imagery into tangible two-dimensional visual representations."],
    ["Paint", "Paint the wall", "Paint a picture", "The famous artist loves to paint beautiful landscapes", "Applying pigmented emulsions is a fundamental technique within classical visual aesthetic disciplines."],
    ["Dance", "I can dance", "Dance with me", "They danced gracefully together at the grand ball", "Choreographed rhythmic kinetics serve as a universal medium for profound emotional expression."],
    ["Skip", "Skip the rope", "Let us skip", "The happy children skip joyfully down the street", "Alternating ambulatory hops develop critical neurological coordination and exceptional lower-body kinetic equilibrium."],
    ["Crawl", "The baby crawls", "Crawl on floor", "The soldier had to crawl under the wire", "Quadrupedal ground traversal drastically minimizes the visible physiological profile against hostile observation."],
    ["Dive", "Dive in water", "Dive very deep", "The brave swimmer will dive into the pool", "Hydrodynamic entry techniques require absolute mastery of mid-air physiological spatial orientation."],
    ["Fly", "The bird flies", "Fly an airplane", "We will fly to Europe for our vacation", "Aerodynamic navigation defies gravitational constraints utilizing sophisticated mechanical or biological lift generation."],
    ["Ride", "Ride a bike", "Ride a horse", "I love to ride my bicycle every weekend", "Equestrian or vehicular manipulation necessitates intuitive kinetic harmony between pilot and apparatus."],
    ["Think", "I can think", "Think very hard", "You must think carefully before making a decision", "Cognitive deliberation synthesizes disparate informational variables to formulate optimal strategic resolutions."],
    ["Work", "I must work", "Work very hard", "My father goes to work early every morning", "Professional labor constitutes the fundamental macroeconomic engine driving continuous global societal progression."],
    ["Study", "I study hard", "Study for test", "We need to study diligently for the exam", "Rigorous academic inquiry systematically expands the boundaries of comprehensive human intellectual paradigms."],
    ["Joyful", "I feel joyful", "A joyful day", "The joyful music made everyone want to dance", "Profound psychological euphoria is correlated with optimal neurochemical dopamine and serotonin saturation."],
    ["Crying", "The baby is crying", "Stop crying now", "The crying child lost his favorite little toy", "Lachrymal secretion is a primary physiological mechanism for alleviating intense emotional distress."],
    ["Mad", "I am mad", "Do not be mad", "He was very mad when he lost game", "Aggressive emotional volatility often stems from profound feelings of systemic injustice or frustration."],
    ["Exhausted", "I am exhausted", "He looks exhausted", "The exhausted runner collapsed after finishing the marathon", "Severe physiological depletion requires immediate and sustained periods of deep restorative recuperation."],
    ["Starving", "I am starving", "The dog is starving", "We are starving because we missed our lunch", "Critical caloric deficits trigger urgent neurological imperatives demanding immediate nutritional systemic replenishment."],
    ["Dry", "The leaf is dry", "A dry towel", "The dry desert has almost no water inside", "Absolute moisture deprivation creates incredibly harsh environments entirely devoid of sustaining biological hydration."],
    ["Huge", "A huge elephant", "The building is huge", "They discovered a huge ancient pyramid hidden away", "Colossal structural dimensions frequently overwhelm typical baseline parameters of standard human spatial comprehension."],
    ["Tiny", "A tiny ant", "The bug is tiny", "I found a tiny little shell on beach", "Microscopic physiological scaling allows organisms to inhabit profoundly intricate and marginalized ecological niches."],
    ["Tall", "A tall tree", "The man is tall", "The tall skyscraper touches the high white clouds", "Vertical architectural elongation maximizes the functional utility of severely limited metropolitan geographic real-estate."],
    ["Short", "A short boy", "The stick is short", "She wore a short dress to the party", "Truncated physical dimensions often necessitate strategic adaptations to navigate environments designed for averages."],
    ["Long", "A long snake", "The road is long", "We endured a very long and exhausting journey", "Extended spatial trajectories frequently demand significantly elevated expenditures of navigational kinetic energy."],
    ["Wide", "A wide river", "The street is wide", "The wide bridge spans across the entire valley", "Broad horizontal parameters facilitate the simultaneous accommodation of massive multifactorial logistical throughputs."],
    ["Narrow", "A narrow path", "The road is narrow", "We squeezed through the dark and narrow alleyway", "Constricted spatial corridors inherently limit the velocity and volume of potential systemic transversal."],
    ["Deep", "A deep hole", "The ocean is deep", "The submarine dived into the deep dark ocean", "Profound bathymetric depths conceal radically alien ecosystems fundamentally adapted to extreme hydrostatic pressures."],
    ["Shallow", "A shallow pool", "The water is shallow", "Children can play safely in the shallow water", "Minimal aquatic depth parameters inherently restrict the viability of sustaining larger marine biological organisms."],
    ["Thick", "A thick book", "The wall is thick", "She read a very thick and complicated dictionary", "Dense material composition provides significantly enhanced resistance against abrasive external kinetic or thermal forces."],
    ["Thin", "A thin paper", "The ice is thin", "Be careful not to break the thin glass", "Flimsy structural integrity demands exceedingly cautious and precise methodological handling to prevent catastrophic fracturing."],
    ["Weighty", "A weighty box", "The stone is weighty", "The weighty responsibility rests entirely on his shoulders", "Massive gravitational resistance necessitates the deployment of specialized mechanical leverage for successful physical displacement."],
    ["Bright", "A bright light", "The star is bright", "The bright sun blinded me for a moment", "Intense photonic emission causes immediate reflexive constriction of the human ocular pupillary mechanism."],
    ["Dark", "A dark room", "The night is dark", "I cannot see anything in this dark cave", "The total absence of ambient illumination fundamentally cripples baseline biological visual navigational capabilities."],
    ["Smooth", "A smooth stone", "The glass is smooth", "The smooth silk fabric feels wonderful against skin", "Frictionless material surfaces reflect meticulous refinement designed to optimize aerodynamic or tactile systemic interactions."],
    ["Rough", "A rough rock", "The road is rough", "The rough sandpaper smoothed out the wooden table", "Abrasive textural irregularities generate significant kinetic friction impeding rapid or fluid superficial traversal."],
    ["Wealthy", "A wealthy king", "The man is wealthy", "The wealthy businessman donated millions to the charity", "Abundant fiscal resources provide unparalleled systemic leverage to unilaterally manipulate complex macroeconomic market paradigms."],
    ["Broke", "I am broke", "The broke student", "He is completely broke and cannot buy food", "Absolute financial insolvency dictates a terrifyingly precarious existence entirely devoid of fundamental socioeconomic security."],
    ["Gorgeous", "A gorgeous dress", "She is gorgeous", "The gorgeous sunset painted the sky in gold", "Exquisite aesthetic harmonization evokes an almost universally profound and transcendent human psychological appreciation."],
    ["Hideous", "A hideous mask", "The monster is hideous", "The hideous architecture completely ruined the beautiful cityscape", "Grotesque visual dissonance actively triggers deeply ingrained biological impulses of visceral psychological revulsion."],
    ["Powerful", "A powerful machine", "The engine is powerful", "The powerful storm destroyed many houses last night", "Overwhelming kinetic output effortlessly obliterates the structural parameters of less robust environmental obstacles."],
    ["Fragile", "A fragile cup", "The glass is fragile", "Please handle this extremely fragile antique with care", "Inherent structural vulnerability absolutely guarantees catastrophic systemic failure upon encountering even minor kinetic stress."],
    ["Bold", "A bold hero", "Be a bold leader", "The bold explorer ventured into the unknown jungle", "Audacious psychological resilience empowers individuals to completely disregard conventional parameters of systemic risk aversion."],
    ["Scared", "I am scared", "Do not be scared", "The little child was scared of the thunder", "Acute psychological terror triggers an overwhelming involuntary adrenaline cascade demanding immediate fight or flight."],
    ["Day", "A sunny day", "The day is bright", "We will go to the park this day", "The diurnal solar cycle dictates the primary operational rhythms of almost all terrestrial biology."],
    ["Night", "A dark night", "The night is cold", "The stars shine beautifully in the dark night", "Nocturnal atmospheric phases facilitate essential physiological recuperation and radically alter baseline ecological predation paradigms."],
    ["Morning", "Good morning", "Wake up this morning", "I drink hot coffee every single morning happily", "The initial auroral phase sets the fundamental psychological and physiological trajectory for the entire cycle."],
    ["Afternoon", "Good afternoon", "Play in the afternoon", "The sun is hottest during the late afternoon", "Post-meridian temporal spans frequently correlate with a noticeable decline in optimal cognitive processing efficiency."],
    ["Evening", "Good evening", "Rest in the evening", "We enjoyed a peaceful walk in the evening", "Crepuscular transitional periods biologically signal the necessary deceleration of intense daily metabolic and kinetic output."],
    ["Week", "One long week", "Seven days a week", "We have a difficult math test next week", "Standardized heptad temporal compartmentalization universally dictates the rhythm of global macroeconomic professional labor cycles."],
    ["Month", "A new month", "January is a month", "My birthday is in the month of July", "Lunar-derived calendar segmentation provides the foundational framework for analyzing medium-term fiscal and agricultural trends."],
    ["Year", "A happy new year", "Twelve months a year", "They celebrated their tenth wedding anniversary this year", "Planetary orbital completion represents the ultimate macro-metric for evaluating profound historical and evolutionary progression."],
    ["Today", "It is today", "Today is a good day", "I have to finish all my homework today", "Immediate present temporal parameters demand absolute priority focus to successfully execute urgent operational protocols."],
    ["Tomorrow", "Wait until tomorrow", "Tomorrow is a new day", "We will travel to the beautiful beach tomorrow", "Anticipatory future planning strictly relies on the predictable continuity of established current systemic variables."],
    ["Yesterday", "It was yesterday", "Yesterday was Sunday", "I visited my grandparents house late yesterday afternoon", "Retrospective temporal analysis is absolutely vital for decoding the causal catalysts of current systemic paradigms."],
    ["Now", "Do it right now", "I am busy now", "You must immediately evacuate the dangerous building now", "Instantaneous action execution entirely bypasses the perilous uncertainties inherently associated with prolonged strategic deliberation."],
    ["Later", "See you later", "I will do it later", "We can discuss this complicated problem much later", "Strategic postponement often serves as a necessary mechanism to prioritize infinitely more critical immediate crises."],
    ["Soon", "Coming very soon", "It will rain soon", "The highly anticipated movie will be released soon", "Imminent temporal proximity significantly escalates the baseline psychological tension surrounding an expected sequential event."],
    ["Never", "I will never", "Never give up", "He promised that he would never lie again", "Absolute temporal negation represents the most extreme and uncompromising parameter of infinite theoretical impossibility."],
    ["Pen", "A blue pen", "Write with a pen", "I need a black pen to sign this", "Ink-dispensing styluses remain the undisputed foundational implements for the permanent transcription of legal documentation."],
    ["Paper", "A white paper", "Draw on the paper", "Please write your answer on the blank paper", "Processed cellulose sheets have historically served as the primary physical medium for intergenerational knowledge transmission."],
    ["Notebook", "A new notebook", "Write in the notebook", "She filled her notebook with brilliant creative ideas", "Bound manuscript compilations provide an optimally organized repository for the systematic archiving of continuous academic data."],
    ["Ruler", "A long ruler", "Measure with a ruler", "Use a ruler to draw a perfectly straight line", "Linear calibration instruments are absolutely indispensable for executing precision architectural and geometrical spatial drafting."],
    ["Eraser", "A pink eraser", "Use the eraser", "He used an eraser to fix his mistake", "Friction-based graphite removal tools offer a critical margin of error within meticulous mathematical and illustrative procedures."],
    ["Bag", "A heavy bag", "Carry your bag", "I packed all my books inside the bag", "Portable containment vessels dramatically exponentially increase the individual logistical capacity for mobile resource transportation."],
    ["Backpack", "A red backpack", "Wear the backpack", "The hiker carried a heavy backpack up mountain", "Dual-strap dorsal containment systems are ergonomically engineered to optimally distribute massive payload gravitational strain."],
    ["Desk", "A wooden desk", "Sit at the desk", "The teacher stood proudly behind the large desk", "Dedicated academic workstations are specifically designed to physically facilitate sustained periods of intense cognitive concentration."],
    ["Board", "A white board", "Look at the board", "The professor wrote the complex formula on board", "Classroom instructional displays serve as the singular focal epicenter for synchronized mass educational visual transmission."],
    ["Chalk", "White chalk", "Write with chalk", "The teacher used chalk to draw a circle", "Compressed calcium carbonate cylinders provide a highly visible and easily erasable medium for temporary academic instruction."],
    ["Marker", "A red marker", "Draw with a marker", "He used a permanent marker to label box", "Felt-tipped chromatic dispensers deploy incredibly vibrant and indelibly saturated pigments upon highly diverse structural surfaces."],
    ["Map", "A world map", "Look at the map", "We used a detailed map to navigate safely", "Cartographic projections abstract incredibly complex geographical topographies into highly simplified and universally comprehensible navigational models."],
    ["Globe", "A round globe", "Spin the globe", "The student pointed to Australia on the globe", "Spherical planetary models offer the only completely undistorted representation of actual global macroeconomic and geological spatial realities."],
    ["Homework", "Do your homework", "Finish the homework", "The difficult homework took three hours to complete", "Extracurricular academic assignments are strategically formulated to ruthlessly reinforce and solidify complex newly introduced cognitive paradigms."],
    ["Test", "A math test", "Study for the test", "She received a perfect score on her test", "Standardized evaluative metrics objectively quantify the precise depth and retention of specific scholastic conceptual mastery."],
    ["Exam", "A final exam", "Pass the exam", "The university entrance exam was incredibly difficult today", "Comprehensive summative assessments absolutely dictate the ultimate trajectory of an individual's subsequent professional and academic career."],
    ["Lesson", "A good lesson", "Learn the lesson", "The history lesson today was very profoundly fascinating", "Structured pedagogical modules systematically dismantle overwhelmingly complex ideologies into easily digestible and sequentially cumulative cognitive components."],
    ["Class", "A loud class", "Go to your class", "The entire class cheered loudly for the winner", "Synchronized educational cohorts foster highly dynamic environments essential for complex collaborative interpersonal intellectual development."],
    ["Grade", "A good grade", "Get a high grade", "Her excellent grade guaranteed her a prestigious scholarship", "Alphanumeric evaluative classifications serve as the universally recognized baseline currency within the ruthless global academic economy."],
    ["Score", "A perfect score", "Check your score", "He achieved the highest possible score on test", "Numerical achievement quantifications relentlessly stratify academic populations based entirely upon objective demonstrations of supreme intellectual competence."],
    ["Goal", "A big goal", "Reach the goal", "He worked hard to achieve his primary goal", "Strategic goal-setting is imperative for sustained long-term institutional success."],
    ["Dream", "A good dream", "Follow your dream", "She chased her dream of becoming a famous doctor", "Visionary dreams frequently serve as the initial catalyst for profound societal transformation."],
    ["Future", "The future", "A bright future", "We must prepare ourselves for the unknown future", "Predictive modeling attempts to forecast the complex trajectories of future macroeconomic trends."],
    ["Success", "Great success", "Achieve your success", "His new business was a huge success this year", "Genuine success demands a relentless synthesis of unwavering dedication and strategic adaptability."],
    ["Winner", "The first winner", "He is the winner", "The proud winner lifted the gold trophy high", "Triumphant victors often dictate the historical narratives shaping subsequent cultural paradigms."]
  ], []);

  // 1. Hitung Siklus Hari berdasarkan Total Data di Bank Kalimat
  // FIX #4 (TIMEZONE): Hitung dayOfYear dari komponen tanggal LOKAL (bukan epoch ms)
  // agar konsisten dengan getTodayDateLocal() dan tidak ada shift ±1 hari di zona non-UTC.
  const maxDays = SENTENCE_BANK.length; // Otomatis menyesuaikan (sekarang 365+)
  const todayDate = new Date();
  // Gunakan new Date(y, m-1, d) — lokal — bukan epoch arithmetic berbasis UTC
  const startOfYear = new Date(todayDate.getFullYear(), 0, 1);
  const todayLocal = new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate());
  const oneDay = 1000 * 60 * 60 * 24;
  const dayOfYear = Math.floor((todayLocal.getTime() - startOfYear.getTime()) / oneDay);
  const currentCycleDay = dayOfYear % maxDays; // Siklus dinamis (Setahun Penuh)

  // 2. Deteksi Role dan Level
  const isStudent = user.role === 'student';
  const todayStr = getTodayDateLocal();
  const student = isStudent ? db.students.find(s => s.id === user.studentId) : null;
  const isCompletedToday = isStudent ? student?.lastSpeakingChallengeDate === todayStr : false;

  // State Pratinjau Admin/Tutor
  const [previewLevel, setPreviewLevel] = useState(0); 
  const [previewDay, setPreviewDay] = useState(currentCycleDay);

  // 3. Tentukan Index Target berdasarkan Role
  let activeLevelIndex = 0;
  if (isStudent && student) {
     const sessionName = getStudentSession(student);
     activeLevelIndex = SESSIONS.indexOf(sessionName);
     if (activeLevelIndex === -1) activeLevelIndex = 0;
  } else {
     activeLevelIndex = previewLevel;
  }

  const activeDay = isStudent ? currentCycleDay : previewDay;
  const safeDayIndex = Math.max(0, Math.min(maxDays - 1, activeDay));
  const safeLevelIndex = Math.max(0, Math.min(4, activeLevelIndex));
  
  const TARGET_SENTENCE = SENTENCE_BANK[safeDayIndex][safeLevelIndex];

  // 4. State Pengenalan Suara
  const [isListening, setIsListening] = useState(false);
  const [spokenText, setSpokenText] = useState('');
  const [speakScore, setSpeakScore] = useState(null);
  const [speechSupport, setSpeechSupport] = useState(true);
  const [isSafariOrIOS, setIsSafariOrIOS] = useState(false);

  useEffect(() => {
     // Cek dukungan API native
     const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
     if (!SpeechRecognition) {
        setSpeechSupport(false);
     }
     
     // Deteksi User Agent untuk iOS atau Safari
     if (typeof navigator !== 'undefined') {
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !(window as any).MSStream;
        const isSafari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent);
        if (isIOS || isSafari) setIsSafariOrIOS(true);
     }

     // FIX #14 (AUDIOCONTEXT LEAK): Tutup AudioContext saat komponen unmount
     // agar tidak terjadi resource leak jika komponen di-mount/unmount berkali-kali.
     return () => {
       if (audioCtx && audioCtx.state !== 'closed') {
         audioCtx.close().catch(() => {});
         audioCtx = null;
       }
     };
  }, []);

  const startSpeakingChallenge = () => {
     if (isCompletedToday && isStudent) {
        return showToast(language === 'id' ? 'Anda sudah menyelesaikan tantangan hari ini. Kembali lagi besok!' : 'You already completed today\'s challenge. Come back tomorrow!', 'warning');
     }

     // WAJIB: Meminta izin audio pada saat interaksi pengguna untuk mencegah blokir autoplay browser
     initAudioContext();

     const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
     if (!SpeechRecognition) {
        setSpeechSupport(false);
        return;
     }

     const recognition = new SpeechRecognition();
     recognition.lang = 'en-US'; 
     recognition.interimResults = false;
     recognition.maxAlternatives = 1;

     recognition.onstart = () => {
        setIsListening(true);
        setSpokenText('');
        setSpeakScore(null);
     };

     recognition.onend = () => {
        setIsListening(false);
     };

     recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript;
        setSpokenText(transcript);
        
        const normalize = (str) => str.toLowerCase().replace(/[^\w\s]|_/g, "").replace(/\s+/g, " ").trim();
        const targetWords = normalize(TARGET_SENTENCE).split(' ');
        const spokenWords = normalize(transcript).split(' ');
        
        let matches = 0;
        targetWords.forEach(w => {
           if(spokenWords.includes(w)) matches++;
        });
        
        const score = Math.round((matches / targetWords.length) * 100);
        const finalScore = Math.min(100, score);
        setSpeakScore(finalScore);

        // Jika berhasil lulus (misal score > 50%)
        if (finalScore >= 50) {
           playFeedbackSound('success');
           if (isStudent) {
              setDb(prev => ({
                 ...prev,
                 students: prev.students.map(s => {
                    if (s.id === user.studentId) {
                       return { 
                          ...s, 
                          lastSpeakingChallengeDate: todayStr,
                          speakingChallengeCompletedCount: (s.speakingChallengeCompletedCount || 0) + 1,
                          updatedAt: getSyncTimestamp(), // BUGFIX (BUG 6): LWW wajib tahu versi ini lebih baru
                       };
                    }
                    return s;
                 })
              }));
              setTimeout(() => {
                 showToast(language === 'id' ? 'Tantangan Selesai! Anda mendapatkan +20 EXP!' : 'Challenge Completed! You got +20 EXP!', 'success');
              }, 500);
           } else {
              setTimeout(() => {
                 showToast(language === 'id' ? 'Uji Coba Berhasil! (Mode Pratinjau)' : 'Test Successful! (Preview Mode)', 'success');
              }, 500);
           }
        } else {
           playFeedbackSound('error');
           setTimeout(() => {
              showToast(language === 'id' ? 'Skor masih kurang. Coba lagi agar lebih jelas!' : 'Score is too low. Try again clearer!', 'warning');
           }, 500);
        }
     };

     recognition.onerror = (e) => {
        setIsListening(false);
        if(e.error === 'not-allowed') {
           showToast("Please allow microphone access to use this feature.", "error");
        }
     };
     
     recognition.start();
  };

  return (
    <div className="space-y-6 animation-fade-in font-sans pb-8 max-w-4xl mx-auto">
      
      {/* PANEL PRATINJAU KHUSUS ADMIN / TUTOR */}
      {!isStudent && (
        <Card className="bg-[#151B26] border-purple-500/30 shadow-[0_4px_20px_rgba(168,85,247,0.1)] mb-4">
           <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                 <h3 className="text-white font-bold flex items-center gap-2"><Eye size={18} className="text-purple-400" /> Admin/Tutor Preview Control</h3>
                 <p className="text-xs text-gray-400 mt-1">Select a level and day to preview the exact sentence students will see.</p>
              </div>
              <div className="flex flex-wrap gap-3 w-full sm:w-auto items-center">
                 <select 
                    className="bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-purple-500"
                    value={previewLevel}
                    onChange={(e) => setPreviewLevel(Number(e.target.value))}
                 >
                    <option value={0}>Level 1 (PAUD/TK)</option>
                    <option value={1}>Level 2 (Grade 1-2)</option>
                    <option value={2}>Level 3 (Grade 3-4)</option>
                    <option value={3}>Level 4 (Grade 5-6)</option>
                    <option value={4}>Level 5 (SMP/SMA+)</option>
                 </select>
                 <div className="flex items-center gap-2 bg-[#0B0F19] border border-gray-700 rounded-lg px-3 py-2">
                    <span className="text-xs text-gray-400 font-bold">Day:</span>
                    <input 
                       type="number" 
                       min="1" 
                       max={maxDays} 
                       className="bg-transparent text-sm text-white focus:outline-none w-16 text-center font-bold"
                       value={previewDay + 1}
                       onChange={(e) => setPreviewDay(Math.max(0, Math.min(maxDays - 1, Number(e.target.value) - 1)))}
                    />
                    <span className="text-xs text-gray-500">/ {maxDays}</span>
                 </div>
              </div>
           </div>
        </Card>
      )}

      <Card className="border-t-4 border-t-[#00D4FF] shadow-[0_8px_30px_rgba(0,212,255,0.1)] p-0 overflow-hidden relative">
         <div className="absolute top-0 right-0 w-64 h-64 bg-[#00D4FF]/5 blur-3xl rounded-full pointer-events-none"></div>
         
         <div className="p-6 border-b border-gray-800 bg-[#0A0E17] flex justify-between items-center relative z-10">
            <div>
               <h3 className="text-xl font-bold text-white flex items-center gap-2">
                  <Mic size={24} className="text-[#00D4FF]" /> {language === 'id' ? 'Tantangan Berbicara Harian' : 'Daily Speaking Challenge'}
               </h3>
               <p className="text-sm text-gray-400 mt-1">
                  {!isStudent 
                    ? `Pratinjau: Level ${previewLevel + 1} - Hari ${previewDay + 1}`
                    : (language === 'id' ? 'Baca kalimat dengan jelas untuk mendapatkan EXP harian!' : 'Read the sentence clearly to earn daily EXP!')}
               </p>
            </div>
            <div className="flex gap-2">
               {!isStudent && <span className="bg-purple-500/20 text-purple-400 px-3 py-1 rounded-full text-xs font-bold">{language === 'id' ? 'PRATINJAU' : 'PREVIEW'}</span>}
               {!speechSupport && <span className="bg-red-500/20 text-red-400 px-3 py-1 rounded-full text-xs font-bold">{language === 'id' ? 'Peramban Tidak Mendukung' : 'Browser Not Supported'}</span>}
            </div>
         </div>

         <div className="p-8 bg-[#151B26] relative z-10 flex flex-col items-center justify-center text-center min-h-[400px]">
            
            {(!speechSupport || isSafariOrIOS) && !isCompletedToday && (
               <div className="w-full max-w-2xl bg-amber-500/10 border border-amber-500/20 text-amber-400 p-4 rounded-xl mb-6 flex items-start gap-3 text-left animation-fade-in shadow-md">
                  <AlertCircle size={24} className="flex-shrink-0" />
                  <p className="text-sm font-medium leading-relaxed">
                     {language === 'id' 
                        ? 'Peramban Safari atau perangkat iOS sering kali membatasi fitur mikrofon bawaan. Gunakan Google Chrome untuk pengalaman terbaik.' 
                        : 'Safari browsers or iOS devices often block native voice features. Please use Google Chrome for the best experience.'}
                  </p>
               </div>
            )}

            {isCompletedToday ? (
               <div className="flex flex-col items-center animation-fade-in">
                  <div className="w-24 h-24 bg-emerald-500/10 rounded-full flex items-center justify-center border border-emerald-500/30 mb-6 shadow-[0_0_30px_rgba(16,185,129,0.2)]">
                     <CheckCircle2 size={48} className="text-emerald-400" />
                  </div>
                  <h3 className="text-2xl font-black text-white mb-2">{language === 'id' ? 'Luar Biasa!' : 'Awesome!'}</h3>
                  <p className="text-gray-400 max-w-sm mb-6">{language === 'id' ? 'Anda sudah menyelesaikan tantangan hari ini dan mendapatkan EXP. Kembali lagi besok untuk misi baru!' : 'You have completed today\'s challenge and earned EXP. Come back tomorrow for a new mission!'}</p>
                  <span className="text-sm font-bold bg-[#0B0F19] border border-gray-700 px-6 py-2 rounded-full text-yellow-400 shadow-inner flex items-center gap-2">
                     <Zap size={16}/> {language === 'id' ? '+20 EXP Tersimpan' : '+20 EXP Saved'}
                  </span>
               </div>
            ) : (
               <>
                  <p className="text-sm text-gray-400 uppercase tracking-widest font-bold mb-4">{language === 'id' ? 'Kalimat Target' : 'Target Sentence'}</p>
                  <div className="bg-[#0B0F19] border border-gray-700 px-8 py-6 rounded-2xl mb-8 shadow-inner max-w-2xl w-full">
                     <p className="text-2xl sm:text-3xl font-black text-white italic">"{TARGET_SENTENCE}"</p>
                  </div>

                  <button 
                     onClick={startSpeakingChallenge}
                     disabled={!speechSupport || isListening}
                     className={`w-20 h-20 sm:w-24 sm:h-24 rounded-full flex items-center justify-center transition-all duration-300 shadow-xl border-4 ${
                        isListening 
                        ? 'bg-rose-500 border-rose-400 animate-pulse scale-110 shadow-[0_0_40px_rgba(244,63,94,0.6)]' 
                        : 'bg-[#00D4FF] border-[#0B0F19] hover:bg-[#00a8cc] hover:scale-105 shadow-[0_0_30px_rgba(0,212,255,0.4)]'
                     } disabled:opacity-50 disabled:cursor-not-allowed`}
                  >
                     <Mic size={40} className={isListening ? "text-white" : "text-[#0B0F19]"} />
                  </button>
                  <p className="mt-4 text-sm font-bold text-gray-400">
                     {isListening ? (language === 'id' ? "Mendengarkan... Silakan bicara sekarang." : "Listening... Please speak now.") : (language === 'id' ? "Ketuk mikrofon dan mulai bicara" : "Tap the mic and start speaking")}
                  </p>

                  {/* RESULTS DISPLAY */}
                  {(spokenText || speakScore !== null) && (
                     <div className="mt-8 w-full max-w-2xl bg-[#0B0F19] border border-gray-700 p-6 rounded-2xl animation-fade-in flex flex-col md:flex-row items-center gap-6 text-left">
                        <div className="flex-1">
                           <p className="text-[11px] text-gray-500 font-bold uppercase tracking-widest mb-2">{language === 'id' ? 'Anda Mengatakan:' : 'You Said:'}</p>
                           <p className="text-lg text-gray-300 font-medium italic">"{spokenText || '...' }"</p>
                        </div>
                        
                        <div className="w-full md:w-auto flex flex-col items-center justify-center p-4 bg-[#151B26] border border-gray-800 rounded-xl shadow-inner min-w-[120px]">
                           <span className="text-[11px] text-gray-500 font-bold uppercase tracking-widest mb-1">{language === 'id' ? 'Akurasi' : 'Accuracy'}</span>
                           <span className={`text-4xl font-black ${
                              speakScore >= 50 ? 'text-emerald-400' : 'text-red-400'
                           }`}>
                              {speakScore}%
                           </span>
                        </div>
                     </div>
                  )}
               </>
            )}
         </div>
      </Card>
    </div>
  );
}

function StudentQuestsModule({ db, user, language = 'en', setActiveTab, downloadPNG, handleShareImage }) {
  const currentMonth = String(new Date().getMonth() + 1).padStart(2, '0');
  const currentYear = String(new Date().getFullYear());
  const monthPrefix = `${currentYear}-${currentMonth}`;

  // 1. OTOMATISASI DATA: Baca riwayat siswa dari database tanpa input admin
  const allMyAtt = db.studentAttendance.filter(a => a.studentId === user.studentId);
  const myAttThisMonth = allMyAtt.filter(a => a.date.startsWith(monthPrefix));
  const presentThisMonth = myAttThisMonth.filter(a => a.status === 'Present').length;
  
  const allMyAssessments = db.assessments.filter(a => a.studentId === user.studentId);
  const latestAss = allMyAssessments.length > 0 ? allMyAssessments.sort((a, b) => {
     if (Number(b.year) !== Number(a.year)) return Number(b.year) - Number(a.year);
     return Number(b.month) - Number(a.month);
  })[0] : null;
  const latestScore = latestAss ? latestAss.average : 0;

  let totalTasksDone = 0;
  let tasksDoneThisMonth = 0;
  (db.materials || []).forEach(m => {
     const mySub = (m.submissions || []).find(s => s.studentId === user.studentId);
     if (mySub) {
        totalTasksDone++;
        if (m.date.startsWith(monthPrefix)) tasksDoneThisMonth++;
     }
  });

  // 2. HITUNG TOTAL EXP OTOMATIS (MENGGUNAKAN HELPER)
  const totalEXP = calculateStudentEXP(user.studentId, db);
  const myStudentRecord = db.students.find(s => s.id === user.studentId);
  const speakingChallengeCount = myStudentRecord?.speakingChallengeCompletedCount || 0;

  // 3. KALKULASI LEVEL DIGITAL (MENGGUNAKAN HELPER)
  const levelInfo = getLevelInfo(totalEXP);
  const LevelIcon = levelInfo.icon;

  // 4. BUAT MISI BULANAN (Dinamo)
  const quests = [
    {
       id: 1, title: language === 'id' ? 'Kehadiran Sempurna' : 'Perfect Attendance', desc: language === 'id' ? 'Hadir minimal 4 kelas bulan ini' : 'Attend at least 4 classes this month',
       target: 4, current: presentThisMonth, reward: 200, icon: UserCheck, color: 'emerald'
    },
    {
       id: 2, title: language === 'id' ? 'Juara Akademik' : 'Academic Champion', desc: language === 'id' ? 'Dapatkan rata-rata nilai 85+ bulan ini' : 'Get an average score of 85+ this month',
       target: 85, current: latestScore, reward: 500, icon: Award, color: 'blue'
    },
    {
       id: 3, title: language === 'id' ? 'Ahli Tugas' : 'Task Master', desc: language === 'id' ? 'Kumpulkan minimal 2 tugas bulan ini' : 'Submit at least 2 learning tasks this month',
       target: 2, current: tasksDoneThisMonth, reward: 150, icon: CheckSquare, color: 'purple'
    }
  ];

  return (
    <div className="space-y-6 animation-fade-in font-sans pb-8">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Misi & Lencana Saya' : 'My Quests & Badges'}</h2>
          <p className="text-sm text-gray-400">{language === 'id' ? 'Selesaikan misi otomatis untuk mendapatkan EXP dan tingkatkan peringkatmu!' : 'Complete automated missions to earn EXP and upgrade your rank!'}</p>
        </div>
        <div className="flex gap-2 w-full sm:w-auto">
           <Button onClick={() => downloadPNG('gamification-card', `${user.name.replace(/\s+/g, '_')}_Rank`)} variant="secondary" icon={Download} className="flex-1 sm:flex-none">
             {language === 'id' ? 'Simpan' : 'Save'}
           </Button>
           <Button onClick={() => handleShareImage('gamification-card', `${user.name.replace(/\s+/g, '_')}_Rank`, `I just reached ${levelInfo.title} with ${totalEXP} EXP at English Club Gresik! 🎉`)} icon={Share2} className="flex-1 sm:flex-none bg-gradient-to-r from-pink-500 to-purple-500 hover:from-pink-600 hover:to-purple-600 border-none text-white shadow-[0_0_20px_rgba(217,70,239,0.3)]">
             {language === 'id' ? 'Pamerkan' : 'Share'}
           </Button>
        </div>
      </div>

      {/* Profil Gamifikasi (Banner Atas) */}
      <div id="gamification-card" className={`relative overflow-hidden rounded-[24px] border ${levelInfo.border} p-6 sm:p-8 bg-gradient-to-br from-[#0B0F19] to-[#151B26] shadow-2xl`}>
         <div className="absolute -top-20 -right-20 w-64 h-64 bg-white/5 rounded-full blur-3xl pointer-events-none"></div>
         
         {/* Watermark for sharing context */}
         <div className="absolute top-4 right-6 flex items-center gap-2 opacity-50">
            <span className="text-[11px] font-bold text-white tracking-widest uppercase">English Club Gresik</span>
         </div>

         <div className="flex flex-col md:flex-row items-center gap-6 relative z-10 mt-2">
            <div className={`w-24 h-24 sm:w-32 sm:h-32 rounded-full ${levelInfo.bg} flex flex-col items-center justify-center border-4 ${levelInfo.border} shadow-[0_0_30px_rgba(255,255,255,0.1)] relative shrink-0`}>
               <LevelIcon size={48} className={`text-transparent bg-clip-text bg-gradient-to-b ${levelInfo.color}`} style={{ color: 'white' }} />
               <span className="absolute -bottom-3 bg-[#0B0F19] border border-gray-700 px-3 py-0.5 rounded-full text-xs font-bold text-white uppercase tracking-widest shadow-md">{language === 'id' ? 'PERINGKAT' : 'RANK'}</span>
            </div>
            
            <div className="flex-1 text-center md:text-left w-full">
               <p className="text-gray-400 font-bold uppercase tracking-widest text-xs mb-1">{user.name} • {language === 'id' ? 'Status Saat Ini' : 'Current Status'}</p>
               <h3 className={`text-3xl sm:text-4xl font-black mb-3 text-transparent bg-clip-text bg-gradient-to-r ${levelInfo.color}`}>
                  {levelInfo.title}
               </h3>
               
               <div className="w-full bg-gray-800/80 rounded-full h-4 overflow-hidden border border-gray-700 shadow-inner mb-2 relative">
                  <div className={`h-full rounded-full bg-gradient-to-r ${levelInfo.color} relative transition-all duration-1000`} style={{ width: `${Math.min(100, (totalEXP % 1000) / 10)}%` }}>
                     <div className="absolute top-0 right-0 bottom-0 left-0 bg-white/20 w-full rounded-full"></div>
                  </div>
               </div>
               <div className="flex justify-between text-xs font-bold">
                  <span className="text-white flex items-center gap-1"><Zap size={12} className="text-yellow-400"/> {totalEXP} {language === 'id' ? 'Total EXP' : 'Total EXP'}</span>
                  <span className="text-gray-500">{language === 'id' ? `Peringkat berikutnya pada ${Math.ceil((totalEXP+1)/1000)*1000} EXP` : `Next Rank at ${Math.ceil((totalEXP+1)/1000)*1000} EXP`}</span>
               </div>
            </div>
         </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
         {/* Misi Bulanan (Auto-Tracking) */}
         <Card className="border-[#00D4FF]/20 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="p-5 border-b border-gray-800 bg-[#0A0E17]">
               <h3 className="text-lg font-bold text-white flex items-center gap-2"><Target size={20} className="text-red-400" /> {language === 'id' ? 'Misi Bulanan' : 'Monthly Quests'} ({MONTHS[Number(currentMonth)-1]})</h3>
            </div>
            <div className="p-5 space-y-5 bg-[#151B26] flex-1">
               {quests.map(q => {
                  const isCompleted = q.current >= q.target;
                  const progressPct = Math.min(100, (q.current / q.target) * 100);
                  
                  return (
                     <div key={q.id} className={`p-4 rounded-xl border transition-all ${isCompleted ? 'bg-emerald-500/10 border-emerald-500/30' : 'bg-[#0B0F19] border-gray-800'}`}>
                        <div className="flex gap-4 items-start">
                           <div className={`p-3 rounded-lg flex-shrink-0 ${isCompleted ? 'bg-emerald-500/20 text-emerald-400' : 'bg-gray-800 text-gray-400'}`}>
                              <q.icon size={24} />
                           </div>
                           <div className="flex-1 w-full">
                              <div className="flex justify-between items-start mb-1">
                                 <h4 className={`font-bold ${isCompleted ? 'text-emerald-400' : 'text-white'}`}>{q.title}</h4>
                                 <span className="text-[11px] font-black bg-yellow-500/20 text-yellow-400 px-2 py-0.5 rounded uppercase tracking-wider">+{q.reward} EXP</span>
                              </div>
                              <p className="text-xs text-gray-400 mb-3">{q.desc}</p>
                              
                              <div className="flex items-center gap-3">
                                 <div className="flex-1 bg-gray-800 rounded-full h-2 overflow-hidden shadow-inner">
                                    <div className={`h-full rounded-full transition-all duration-1000 ${isCompleted ? 'bg-emerald-400' : 'bg-[#00D4FF]'}`} style={{ width: `${progressPct}%` }}></div>
                                 </div>
                                 <span className={`text-xs font-bold ${isCompleted ? 'text-emerald-400' : 'text-gray-400'}`}>
                                    {isCompleted ? (language === 'id' ? 'SELESAI' : 'DONE') : `${q.current} / ${q.target}`}
                                 </span>
                              </div>
                           </div>
                        </div>
                     </div>
                  );
               })}
            </div>
         </Card>

         {/* Lencana Pencapaian (Badges) */}
         <Card className="border-purple-500/20 shadow-lg p-0 overflow-hidden flex flex-col">
            <div className="p-5 border-b border-gray-800 bg-[#0A0E17]">
               <h3 className="text-lg font-bold text-white flex items-center gap-2"><Medal size={20} className="text-purple-400" /> {language === 'id' ? 'Lencana Terbuka' : 'Unlocked Badges'}</h3>
            </div>
            <div className="p-6 bg-[#151B26] flex-1 grid grid-cols-3 gap-4">
               {/* Absensi Sempurna */}

               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${presentThisMonth >= 4 ? 'bg-emerald-500/10 border-emerald-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <UserCheck size={32} className={presentThisMonth >= 4 ? 'text-emerald-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">{language === 'id' ? 'Siswa Setia' : 'Loyal Learner'}</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? 'Hadir 4x' : 'Attend 4x'}</span>
               </div>
               
               {/* Nilai Tinggi */}
               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${latestScore >= 90 ? 'bg-blue-500/10 border-blue-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <Award size={32} className={latestScore >= 90 ? 'text-blue-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">{language === 'id' ? 'Pikiran Elit' : 'Elite Mind'}</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? 'Nilai 90+' : 'Score 90+'}</span>
               </div>

               {/* Rajin Mengerjakan Tugas */}
               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${totalTasksDone >= 5 ? 'bg-purple-500/10 border-purple-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <BookOpen size={32} className={totalTasksDone >= 5 ? 'text-purple-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">{language === 'id' ? 'Ahli Tugas' : 'Task Master'}</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? 'Kerjakan 5 Tugas' : 'Do 5 Tasks'}</span>
               </div>
               
               {/* Veteran Member */}
               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${allMyAtt.length >= 20 ? 'bg-amber-500/10 border-amber-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <Star size={32} className={allMyAtt.length >= 20 ? 'text-amber-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">Veteran</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? '20+ Kelas' : '20+ Classes'}</span>
               </div>

               {/* Jagoan Speaking Challenge */}
               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${speakingChallengeCount >= 10 ? 'bg-cyan-500/10 border-cyan-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <Mic size={32} className={speakingChallengeCount >= 10 ? 'text-cyan-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">{language === 'id' ? 'Jagoan Bicara' : 'Speaking Star'}</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? '10+ Tantangan' : '10+ Challenges'}</span>
               </div>

               {/* Legenda EXP */}
               <div className={`flex flex-col items-center justify-center text-center p-3 rounded-xl border ${totalEXP >= 1000 ? 'bg-rose-500/10 border-rose-500/30' : 'bg-gray-800/30 border-gray-800 grayscale opacity-50'}`}>
                  <Trophy size={32} className={totalEXP >= 1000 ? 'text-rose-400 mb-2' : 'text-gray-500 mb-2'} />
                  <span className="text-[11px] font-bold text-white uppercase tracking-wider mb-1 leading-tight">{language === 'id' ? 'Legenda EXP' : 'EXP Legend'}</span>
                  <span className="text-[11px] text-gray-500">{language === 'id' ? '1000+ EXP' : '1000+ EXP'}</span>
               </div>
            </div>
         </Card>
      </div>
      
      {/* Shortcut ke Daily Challenge Module */}
      <Card className="border-[#00D4FF]/20 shadow-lg bg-gradient-to-r from-[#0A3D91] to-[#051126] p-6 sm:p-8 flex flex-col md:flex-row items-center justify-between gap-6 cursor-pointer hover:border-[#00D4FF]/50 transition-all hover:scale-[1.01]" onClick={() => setActiveTab('speaking_challenge')}>
         <div className="flex items-center gap-5">
            <div className="w-16 h-16 rounded-full bg-[#00D4FF]/20 flex items-center justify-center border border-[#00D4FF]/40 shadow-[0_0_20px_rgba(0,212,255,0.3)] shrink-0">
               <Mic size={32} className="text-[#00D4FF]" />
            </div>
            <div>
               <h3 className="text-xl font-black text-white mb-1">{language === 'id' ? 'Tantangan Harian Tersedia!' : 'Daily Challenge Available!'}</h3>
               <p className="text-sm text-blue-200/80">{language === 'id' ? 'Pergi ke menu Tantangan Berbicara Harian untuk mendapatkan +20 EXP setiap harinya.' : 'Go to the Daily Speaking Challenge menu to earn +20 EXP every day.'}</p>
            </div>
         </div>
         <Button className="shrink-0 bg-[#00D4FF] text-[#0B0F19] border-none font-bold px-6 py-3 shadow-[0_0_20px_rgba(0,212,255,0.4)] whitespace-nowrap">
            {language === 'id' ? 'Mulai Sekarang' : 'Start Now'}
         </Button>
      </Card>
    </div>
  );
}

// --- Company letterhead constants used in the Excel export ---
// NOTE: edit COMPANY_ADDRESS to your exact address if this placeholder isn't right.
const COMPANY_NAME = 'ENGLISH CLUB GRESIK';
const COMPANY_TAGLINE = 'English Language Education Center';
const COMPANY_ADDRESS = 'Perumahan Taman Anggrek, Blok AB No. 05, Kedanyang, Kebomas, Gresik, Jawa Timur';
const COMPANY_WA = '0897-327-11-12';
const COMPANY_WEBSITE = 'englishclubgresik.com';
const BRAND_HEX = '00A8CC'; // solid Excel-safe variant of the app's #00D4FF accent
const BRAND_DARK_HEX = '0B0F19';

function DataExportModule({ db }) {
  const exportToExcel = async () => {
    if (!window.ExcelJS) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
      document.body.appendChild(script);
      await new Promise((resolve, reject) => {
        script.onload = resolve;
        script.onerror = reject;
      });
    }
    const ExcelJS = window.ExcelJS;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ECG Smart Portal';
    wb.created = new Date();

    const today = getTodayDateLocal();
    const currentMonth = String(new Date().getMonth() + 1);
    const currentYear = String(new Date().getFullYear());
    const printedAt = new Date().toLocaleString('id-ID', { dateStyle: 'full', timeStyle: 'short' });

    const THIN = { style: 'thin', color: { argb: 'FFB0B7C3' } };
    const ALL_BORDERS = { top: THIN, left: THIN, bottom: THIN, right: THIN };

    // Draws the "kop surat" (company info + report title) at the top of a
    // sheet and returns the row number where the data table should start.
    const addLetterhead = (ws, reportTitle, lastCol) => {
      ws.mergeCells(1, 1, 1, lastCol);
      const nameCell = ws.getCell(1, 1);
      nameCell.value = COMPANY_NAME;
      nameCell.font = { bold: true, size: 16, color: { argb: 'FF' + BRAND_DARK_HEX } };
      nameCell.alignment = { vertical: 'middle' };

      ws.mergeCells(2, 1, 2, lastCol);
      const taglineCell = ws.getCell(2, 1);
      taglineCell.value = COMPANY_TAGLINE;
      taglineCell.font = { italic: true, size: 10, color: { argb: 'FF6B7280' } };

      ws.mergeCells(3, 1, 3, lastCol);
      const addrCell = ws.getCell(3, 1);
      addrCell.value = `${COMPANY_ADDRESS}  |  ${COMPANY_WEBSITE}`;
      addrCell.font = { size: 10, color: { argb: 'FF374151' } };

      ws.mergeCells(4, 1, 4, lastCol);
      const waCell = ws.getCell(4, 1);
      waCell.value = `WhatsApp / Telp: ${COMPANY_WA}`;
      waCell.font = { size: 10, color: { argb: 'FF374151' } };

      ws.getRow(1).height = 20;
      ws.getRow(2).height = 16;
      ws.getRow(3).height = 16;
      ws.getRow(4).height = 16;

      ws.mergeCells(5, 1, 5, lastCol);
      ws.getRow(5).height = 3;
      ws.getCell(5, 1).border = { bottom: { style: 'medium', color: { argb: 'FF' + BRAND_HEX } } };

      ws.mergeCells(6, 1, 6, lastCol);
      const titleCell = ws.getCell(6, 1);
      titleCell.value = reportTitle;
      titleCell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
      titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + BRAND_DARK_HEX } };
      ws.getRow(6).height = 22;

      ws.mergeCells(7, 1, 7, lastCol);
      const metaCell = ws.getCell(7, 1);
      metaCell.value = `Print Date: ${today}   |   Printed At: ${printedAt}   |   Generated By: ECG Smart Portal System`;
      metaCell.font = { size: 9, italic: true, color: { argb: 'FF6B7280' } };
      ws.getRow(7).height = 14;

      return 9; // first free row after a blank spacer row
    };

    // Renders a header row + data rows starting at startRow, with borders,
    // banded rows, autofilter and sensible column widths.
    const addTable = (ws, startRow, headers, rows, options: { centerCols?: number[]; currencyCols?: number[] } = {}) => {
      const headerRow = ws.getRow(startRow);
      headers.forEach((h, i) => {
        const cell = headerRow.getCell(i + 1);
        cell.value = h;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + BRAND_HEX } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = ALL_BORDERS;
      });
      headerRow.height = 20;

      rows.forEach((r, rIdx) => {
        const row = ws.getRow(startRow + 1 + rIdx);
        r.forEach((val, cIdx) => {
          const cell = row.getCell(cIdx + 1);
          cell.value = val;
          cell.border = ALL_BORDERS;
          cell.font = { size: 10 };
          cell.alignment = { vertical: 'middle', horizontal: options.centerCols?.includes(cIdx) ? 'center' : 'left' };
          if (rIdx % 2 === 1) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F6F9' } };
          }
          if (options.currencyCols?.includes(cIdx)) {
            cell.numFmt = '"Rp" #,##0';
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          }
        });
      });

      ws.autoFilter = {
        from: { row: startRow, column: 1 },
        to: { row: startRow, column: headers.length },
      };
      ws.views = [{ state: 'frozen', ySplit: startRow, xSplit: 0 }];

      headers.forEach((h, i) => {
        const maxLen = Math.max(
          String(h).length,
          ...rows.map((r) => (r[i] === null || r[i] === undefined ? 0 : String(r[i]).length))
        );
        ws.getColumn(i + 1).width = Math.min(Math.max(maxLen + 3, 12), 40);
      });

      return startRow + 1 + rows.length;
    };

    // ---------------- Sheet 1: Executive Summary ----------------
    const activeStudentsCount = db.students.filter((s) => s.status === 'Active').length;
    const totalStudentsCount = db.students.length;
    const activeTutorsCount = db.tutors.filter((t) => t.status === 'Active').length;
    const monthlyRevenue = db.payments
      .filter((p) => p.status === 'Paid' && Number(p.month) === Number(currentMonth) && String(p.year) === String(currentYear))
      .reduce((sum, p) => sum + Number(p.amount), 0);
    const totalRevenueAllTime = db.payments
      .filter((p) => p.status === 'Paid')
      .reduce((sum, p) => sum + Number(p.amount), 0);
    const totalAssessments = db.assessments.length;
    const avgScoreAll = totalAssessments
      ? Math.round(db.assessments.reduce((sum, a) => sum + Number(a.average || 0), 0) / totalAssessments)
      : 0;
    const attendanceRateAll = db.studentAttendance.length
      ? Math.round(
          (db.studentAttendance.filter((a) => a.status === 'Present').length / db.studentAttendance.length) * 100
        )
      : 0;

    const wsSummary = wb.addWorksheet('Executive Summary');
    let r = addLetterhead(wsSummary, 'EXECUTIVE SUMMARY', 4);
    r = addTable(
      wsSummary,
      r,
      ['Metric', 'Value'],
      [
        ['Total Registered Students', totalStudentsCount],
        ['Total Active Students', activeStudentsCount],
        ['Total Active Tutors', activeTutorsCount],
        ['Avg Student Attendance (All Time)', `${attendanceRateAll}%`],
        ['Avg Academic Score (All Time)', avgScoreAll],
        ['Current Month Revenue', `Rp ${monthlyRevenue.toLocaleString('id-ID')}`],
        ['Total Revenue (All Time, Paid)', `Rp ${totalRevenueAllTime.toLocaleString('id-ID')}`],
      ],
      { currencyCols: [] }
    );

    // ---------------- Sheet 2: Students Directory ----------------
    const wsStudents = wb.addWorksheet('Students');
    let r2 = addLetterhead(wsStudents, 'STUDENTS DIRECTORY REPORT', 8);
    addTable(
      wsStudents,
      r2,
      ['ID', 'Name', 'Level', 'Class', 'Session', 'WhatsApp No.', 'Payment Plan', 'Status'],
      db.students.map((s) => [
        s.id,
        s.name,
        s.level || '',
        s.class || '',
        getStudentSession(s),
        s.whatsapp ? String(s.whatsapp).replace(/^'/, '') : '',
        s.paymentPlan || '',
        s.status,
      ]),
      { centerCols: [4, 7] }
    );

    // ---------------- Sheet 3: Tutors Directory ----------------
    const wsTutors = wb.addWorksheet('Tutors');
    let r3 = addLetterhead(wsTutors, 'TUTORS DIRECTORY REPORT', 6);
    addTable(
      wsTutors,
      r3,
      ['ID', 'Name', 'Specialization / Teaching Session', 'Phone / WA No.', 'Status', 'Total Sessions Attended'],
      db.tutors.map((t) => {
        const totalSessions = db.tutorAttendance.filter((a) => a.tutorId === t.id && a.status === 'Present').length;
        return [t.id, t.name, parseSessions(t.teachingSession).join(', ') || '', t.phone || '', t.status || '', totalSessions];
      }),
      { centerCols: [4, 5] }
    );

    // ---------------- Sheet 4: Student Attendance Summary ----------------
    const wsAttendance = wb.addWorksheet('Attendance Summary');
    let r4 = addLetterhead(wsAttendance, 'STUDENT ATTENDANCE SUMMARY', 6);
    const attendanceByStudent = db.students.map((s) => {
      const recs = db.studentAttendance.filter((a) => a.studentId === s.id);
      const present = recs.filter((a) => a.status === 'Present').length;
      const total = recs.length;
      const rate = total ? Math.round((present / total) * 100) : 0;
      return [s.id, s.name, s.class || '', total, present, `${rate}%`];
    });
    addTable(
      wsAttendance,
      r4,
      ['ID', 'Student Name', 'Class', 'Total Recorded Sessions', 'Total Present', 'Attendance Rate'],
      attendanceByStudent,
      { centerCols: [3, 4, 5] }
    );

    // ---------------- Sheet 5: Academic Performance ----------------
    const wsAcademic = wb.addWorksheet('Academic Performance');
    let r5 = addLetterhead(wsAcademic, 'ACADEMIC PERFORMANCE REPORT', 6);
    const academicRows = db.assessments
      .slice()
      .sort((a, b) => `${b.year}${b.month}`.localeCompare(`${a.year}${a.month}`))
      .map((a) => [
        a.studentId,
        a.studentName,
        a.class || '',
        `${a.month}/${a.year}`,
        a.average,
        getPerformanceCat(a.average).label,
      ]);
    addTable(
      wsAcademic,
      r5,
      ['Student ID', 'Student Name', 'Class', 'Month/Year', 'Average Score', 'Category'],
      academicRows,
      { centerCols: [3, 4, 5] }
    );

    // ---------------- Sheet 6: Financial - Payments ----------------
    const wsPayments = wb.addWorksheet('Financial - Payments');
    let r6 = addLetterhead(wsPayments, 'FINANCIAL REPORT - PAYMENTS', 7);
    const sortedPayments = [...db.payments].sort((a, b) => {
      if (a.date === b.date) return (b.time || '').localeCompare(a.time || '');
      return b.date.localeCompare(a.date);
    });
    const paidTotal = sortedPayments
      .filter((p) => p.status === 'Paid')
      .reduce((sum, p) => sum + Number(p.amount), 0);
    const lastDataRow = addTable(
      wsPayments,
      r6,
      ['Transaction ID', 'Date', 'Student Name', 'Month/Year', 'Class', 'Method', 'Amount (Rp)'],
      sortedPayments.map((p) => [p.id, p.date, p.studentName, `${p.month}/${p.year}`, p.class, p.method || '', Number(p.amount)]),
      { centerCols: [1, 3, 4, 5], currencyCols: [6] }
    );
    const totalRow = wsPayments.getRow(lastDataRow + 1);
    wsPayments.mergeCells(lastDataRow + 1, 1, lastDataRow + 1, 6);
    const totalLabelCell = totalRow.getCell(1);
    totalLabelCell.value = 'TOTAL REVENUE (PAID)';
    totalLabelCell.font = { bold: true, size: 10 };
    totalLabelCell.alignment = { horizontal: 'right' };
    totalLabelCell.border = ALL_BORDERS;
    const totalValueCell = totalRow.getCell(7);
    totalValueCell.value = paidTotal;
    totalValueCell.numFmt = '"Rp" #,##0';
    totalValueCell.font = { bold: true, size: 10, color: { argb: 'FF' + BRAND_DARK_HEX } };
    totalValueCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F9FC' } };
    totalValueCell.border = ALL_BORDERS;
    totalRow.height = 20;

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ECG_Corporate_Report_${today}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-2xl font-bold text-white mb-1">Data Export</h2>
        <p className="text-gray-400 text-sm">Download all application data in corporate report format (.xlsx).</p>
      </div>
      <Card className="flex flex-col items-center justify-center text-center py-16">
        <div className="w-16 h-16 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/30 flex items-center justify-center mb-5">
          <Database size={28} className="text-[#00D4FF]" />
        </div>
        <h3 className="text-lg font-semibold text-white mb-1">Corporate Data Export</h3>
        <p className="text-gray-400 text-sm mb-6 max-w-md">
          Full export with letterhead (logo, address, WhatsApp) covering Executive Summary, Students Directory, Tutors Directory, Attendance Summary, Academic Performance, and Financial Report in a single Excel file.
        </p>
        <button
          onClick={exportToExcel}
          className="flex items-center gap-2 bg-[#00D4FF] hover:bg-[#00B8E0] text-[#0A0E17] font-bold px-6 py-3.5 rounded-xl transition-colors shadow-[0_0_20px_rgba(0,212,255,0.3)]"
        >
          <Download size={20} />
          Download Corporate Report (.xlsx)
        </button>
      </Card>
    </div>
  );
}

function RecycleBinModule({ db, setDb, showToast, requestConfirm }) {
  const [selectedIds, setSelectedIds] = useState([]);
  const binItems = db.recycleBin || [];

  // NEW: State for Pagination
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [rowsPerPage]);

  const toggleSelectOne = (binId) => {
    setSelectedIds((prev) => prev.includes(binId) ? prev.filter((id) => id !== binId) : [...prev, binId]);
  };

  const toggleSelectAll = () => {
    setSelectedIds(selectedIds.length === binItems.length ? [] : binItems.map((b) => b.binId));
  };

  const handleRestore = (item) => {
    requestConfirm('Restore Item', 'Are you sure you want to restore this record?', () => {
      // BUGFIX #12: Normalize dates in restored item so ISO strings don't persist
      const restoreData = (() => {
        const d = { ...item.data };
        const toLocalDate = (s) => {
          if (typeof s === 'string' && s.includes('T') && s.includes('Z')) {
            const dt = new Date(s);
            if (!isNaN(dt.getTime())) return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth()+1).padStart(2,'0')}-${String(dt.getUTCDate()).padStart(2,'0')}`;
          }
          return s;
        };
        if (d.date) d.date = toLocalDate(d.date);
        if (d.joinedDate) d.joinedDate = toLocalDate(d.joinedDate);
        if (d.whatsapp) d.whatsapp = normalizeWhatsapp(d.whatsapp);
        if (d.phone) d.phone = normalizeWhatsapp(d.phone);
        return d;
      })();
      setDb((p) => ({ ...p, [item.originalCollection]: [...p[item.originalCollection], restoreData], recycleBin: p.recycleBin.filter((x) => x.binId !== item.binId) }));
      showToast('Item Restored');
    });
  };

  const handlePermDelete = (binId) => {
    requestConfirm('Permanent Delete', 'WARNING: This will permanently delete the record. This cannot be undone. Continue?', () => {
      setDb((p) => ({ ...p, recycleBin: p.recycleBin.filter((x) => x.binId !== binId) }));
      setSelectedIds((prev) => prev.filter((id) => id !== binId));
      showToast('Permanently Deleted', 'error');
    });
  };

  const handleBulkDelete = () => {
    if (selectedIds.length === 0) return;
    requestConfirm(
      'Permanent Delete Selected',
      `WARNING: This will permanently delete ${selectedIds.length} selected record(s). This cannot be undone. Continue?`,
      () => {
        setDb((p) => ({ ...p, recycleBin: p.recycleBin.filter((x) => !selectedIds.includes(x.binId)) }));
        showToast(`${selectedIds.length} item(s) permanently deleted`, 'error');
        setSelectedIds([]);
      }
    );
  };

  const reversedBinItems = [...binItems].reverse();

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? reversedBinItems.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(reversedBinItems.length / (rowsNum || 1));
  const paginatedData = isAll ? reversedBinItems : reversedBinItems.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white mb-1">Recycle Bin</h2><p className="text-gray-400 text-sm">Restore or permanently delete removed records.</p></div>
        <Button onClick={() => requestConfirm('Empty Recycle Bin', 'Are you absolutely sure? This will PERMANENTLY delete all items in the bin. This cannot be undone.', () => { setDb(p => ({ ...p, recycleBin: [] })); showToast('Recycle bin emptied', 'warning'); })} variant="secondary" className="border-red-500/30 text-red-400 hover:bg-red-500 hover:text-white" icon={Trash2} disabled={db.recycleBin.length === 0}>Empty Bin</Button>
      </div>
      
      <Card className="p-0 flex flex-col overflow-hidden">
        <div className="p-4 bg-[#0A0E17] border-b border-gray-800 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-2">
            <ArchiveRestore className="text-yellow-500" />
            <h3 className="font-semibold text-white">Deleted Items</h3>
          </div>
          {binItems.length > 0 && (
            <Button
              variant="danger"
              className="text-xs px-3 py-2"
              icon={Trash}
              disabled={selectedIds.length === 0}
              onClick={handleBulkDelete}
            >
              Delete Selected ({selectedIds.length})
            </Button>
          )}
        </div>
        
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold">
              <tr>
                <th className="p-4 text-center w-10">
                  <input
                    type="checkbox"
                    className="w-4 h-4 cursor-pointer accent-[#00D4FF]"
                    checked={binItems.length > 0 && selectedIds.length === binItems.length}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="p-4 text-center w-12 text-gray-400">No.</th>
                <th className="p-4 text-center">Deleted At</th>
                <th className="p-4 text-center">Module Type</th>
                <th className="p-4 text-center">Data Summary</th>
                <th className="p-4 text-center">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-800">
              {paginatedData.map((b, index) => (
                <tr key={b.binId} className={`hover:bg-[#0B0F19] ${selectedIds.includes(b.binId) ? 'bg-[#00D4FF]/5' : ''}`}>
                  <td className="p-4 text-center">
                    <input
                      type="checkbox"
                      className="w-4 h-4 cursor-pointer accent-[#00D4FF]"
                      checked={selectedIds.includes(b.binId)}
                      onChange={() => toggleSelectOne(b.binId)}
                    />
                  </td>
                  <td className="p-4 text-center text-gray-500 font-medium">{startIndex + index + 1}</td>
                  <td className="p-4 text-center">{normalizeTimestamp(b.deletedAt)}</td>
                  <td className="p-4 text-center capitalize text-[#00D4FF] font-medium">{b.originalCollection}</td>
                  <td className="p-4 text-center text-gray-400 max-w-xs truncate">{JSON.stringify(b.data)}</td>
                  <td className="p-4 text-center flex justify-center gap-2">
                    <button onClick={() => handleRestore(b)} className="text-emerald-400 p-2.5 hover:bg-emerald-500/10 rounded-lg transition-colors" title="Restore Item"><RefreshCw size={18} /></button>
                    <button onClick={() => handlePermDelete(b.binId)} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Permanently"><Trash size={18} /></button>
                  </td>
                </tr>
              ))}
              {paginatedData.length === 0 && (
                <tr><td colSpan={6}><EmptyState icon={ArchiveRestore} title="Recycle bin is empty" description="Deleted items will appear here." /></td></tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer */}
        {binItems.length > 0 && (
          <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
              <div className="flex items-center gap-2">
                <span>Show</span>
                <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                  <option value={10}>10</option>
                  <option value={20}>20</option>
                  <option value={50}>50</option>
                  <option value="All">All</option>
                </select>
                <span>entries (Total: {binItems.length})</span>
              </div>
              
              {!isAll && totalPages > 1 && (
                <div className="flex items-center gap-2">
                  <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                  <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                  <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
                </div>
              )}
          </div>
        )}
      </Card>
    </div>
  );
}

function SystemLogsModule({ logs, setLogs, showToast, requestConfirm }) {
  const [activeLogTab, setActiveLogTab] = useState('audit');
  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearchTerm = useDebounce(searchTerm, 300);

  // NEW: State for Pagination
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(20);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [activeLogTab, debouncedSearchTerm, rowsPerPage]);

  const auditLogs = logs.auditLogs || [];
  const debugLogs = logs.debugLogs || [];

  const currentLogs = activeLogTab === 'debug' ? debugLogs : auditLogs;
  
  const filteredLogs = currentLogs.filter(log => {
     const searchStr = debouncedSearchTerm.toLowerCase();
     return (
        (log.User || '').toLowerCase().includes(searchStr) ||
        (log.Action || '').toLowerCase().includes(searchStr) ||
        (log.Details || log['Error Details'] || '').toLowerCase().includes(searchStr)
     );
  });

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? filteredLogs.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(filteredLogs.length / (rowsNum || 1));
  const paginatedData = isAll ? filteredLogs : filteredLogs.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);
  const startIndex = isAll ? 0 : (currentPage - 1) * rowsNum;

  const clearLogs = () => {
     requestConfirm(
        'Clear Local Logs',
        'Are you sure you want to clear these logs from the application view? (This will not delete them from Google Sheets until fully synced, but it cleans your local dashboard).',
        () => {
           if (activeLogTab === 'debug') {
              setLogs(p => ({ ...p, debugLogs: [] }));
           } else {
              setLogs(p => ({ ...p, auditLogs: [] }));
           }
           showToast(`${activeLogTab === 'debug' ? 'Debug' : 'Audit'} logs cleared locally.`);
        }
     );
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
           <h2 className="text-2xl font-bold text-white mb-1">System Logs Monitor</h2>
           <p className="text-sm text-gray-400">Track application errors and user activities in real-time.</p>
        </div>
        <div className="flex gap-2 w-full sm:w-auto">
           <Button onClick={() => setActiveLogTab('audit')} variant={activeLogTab === 'audit' ? 'primary' : 'secondary'} className={activeLogTab === 'audit' ? 'bg-emerald-500 hover:bg-emerald-600 text-white shadow-[0_0_15px_rgba(16,185,129,0.3)]' : ''}>
              Audit Logs (Activity)
           </Button>
           <Button onClick={() => setActiveLogTab('debug')} variant={activeLogTab === 'debug' ? 'primary' : 'secondary'} className={activeLogTab === 'debug' ? 'bg-red-500 hover:bg-red-600 text-white shadow-[0_0_15px_rgba(239,68,68,0.3)]' : ''}>
              Debug Logs (Errors)
           </Button>
        </div>
      </div>

      <Card className="p-0 overflow-hidden border-t-4 shadow-xl flex flex-col h-[75vh] border-t-gray-700">
         <div className="p-4 sm:p-5 bg-[#0A0E17] border-b border-gray-800 flex flex-col sm:flex-row justify-between items-stretch sm:items-center gap-3 flex-wrap">
            <input 
               type="text" 
               placeholder="Search logs by user, action, or details..." 
               className="flex-1 min-w-[200px] bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-[#00D4FF] transition-all"
               value={searchTerm}
               onChange={e => setSearchTerm(e.target.value)}
            />
            <Button variant="danger" onClick={clearLogs} className="text-xs px-4" icon={Trash2}>Clear View</Button>
         </div>
         
         <div className="flex-1 overflow-auto custom-scrollbar bg-[#0B0F19]">
            <table className="w-full text-left text-sm whitespace-nowrap">
              <thead className="bg-[#0B0F19] border-b border-gray-800 text-gray-400 uppercase tracking-wider text-[11px] font-bold sticky top-0 z-10 shadow-sm">
                <tr>
                  <th className="p-4 text-center w-12 text-gray-400 font-bold uppercase tracking-wider text-[11px]">No.</th>
                  <th className="p-4 text-gray-400 font-bold uppercase tracking-wider text-[11px]">Timestamp</th>
                  <th className="p-4 text-gray-400 font-bold uppercase tracking-wider text-[11px]">User</th>
                  <th className="p-4 text-gray-400 font-bold uppercase tracking-wider text-[11px]">Action</th>
                  <th className="p-4 text-center text-gray-400 font-bold uppercase tracking-wider text-[11px]">Status</th>
                  <th className="p-4 text-gray-400 font-bold uppercase tracking-wider text-[11px] w-1/2">Details</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-800/50">
                 {paginatedData.map((log, idx) => {
                    let displayTime = normalizeTimestamp(log.Timestamp);
                    return (
                    <tr key={idx} className="hover:bg-[#1A2234] transition-colors font-mono text-[13px]">
                       <td className="p-4 text-center text-gray-500 font-medium">{startIndex + idx + 1}</td>
                       <td className="p-4 text-gray-400">{displayTime}</td>
                       <td className="p-4 font-bold text-gray-200">{log.User}</td>
                       <td className="p-4 text-blue-400 font-bold">{log.Action}</td>
                       <td className="p-4 text-center">
                          <span className={`px-2 py-1 rounded text-[11px] font-black uppercase tracking-widest ${
                             log.Status === 'SUCCESS' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 
                             log.Status === 'FAILED' ? 'bg-red-500/10 text-red-400 border border-red-500/20' : 
                             'bg-yellow-500/10 text-yellow-400 border border-yellow-500/20'
                          }`}>
                             {log.Status}
                          </span>
                       </td>
                       <td className="p-4 text-gray-300 whitespace-normal min-w-[300px]">
                          {log.Details || log['Error Details']}
                       </td>
                    </tr>
                 )})}
                 {paginatedData.length === 0 && (
                    <tr>
                       <td colSpan={6} className="p-12 text-center text-gray-500 bg-[#0B0F19] font-sans">
                          <Terminal size={48} className="mx-auto mb-4 text-gray-700" />
                          <p className="font-bold text-lg text-white mb-1">No logs found</p>
                          <p className="text-sm">There are no {activeLogTab} logs matching your criteria.</p>
                       </td>
                    </tr>
                 )}
              </tbody>
            </table>
         </div>

         {/* Pagination Footer */}
         <div className="p-4 bg-[#0A0E17] border-t border-gray-800 flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400">
             <div className="flex items-center gap-2">
               <span>Show</span>
               <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                 <option value={20}>20</option>
                 <option value={50}>50</option>
                 <option value={100}>100</option>
                 <option value="All">All</option>
               </select>
               <span>entries {filteredLogs.length > 0 && `(Total: ${filteredLogs.length})`}</span>
             </div>
             
             {!isAll && totalPages > 1 && (
               <div className="flex items-center gap-2">
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                 <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
               </div>
             )}
         </div>
      </Card>
    </div>
  );
}

const getLinkPreview = (url) => {
   if(!url) return { type: 'none', src: null, videoId: null };
   try {
       // Cek apakah ini link YouTube
       const ytRegExp = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|\&v=)([^#\&\?]*).*/;
       const ytMatch = url.match(ytRegExp);
       if (ytMatch && ytMatch[2].length === 11) {
          return { type: 'youtube', src: `https://img.youtube.com/vi/${ytMatch[2]}/hqdefault.jpg`, videoId: ytMatch[2] };
       }
       // Jika bukan YouTube, gunakan thum.io untuk auto-screenshot website
       return { type: 'website', src: `https://image.thum.io/get/width/600/crop/600/noanimate/${url}`, videoId: null };
   } catch(e) {
       return { type: 'none', src: null, videoId: null };
   }
};

function MaterialsModule({ db, setDb, generateId, showToast, softDelete, user }) {
  const defaultSession = user.role === 'tutor' ? (parseSessions(user.teachingSession)[0] || SESSIONS[0]) : SESSIONS[0];
  const [formData, setFormData] = useState({ title: '', sessionGroup: defaultSession, link: '', notes: '' });
  const [isAdding, setIsAdding] = useState(false);
  const [isEditingId, setIsEditingId] = useState(null);
  const [replyTexts, setReplyTexts] = useState({});
  const [playingVideos, setPlayingVideos] = useState({});
  
  // NEW: State for Filters & Pagination
  const [filterMonth, setFilterMonth] = useState<number | string>(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  const [filterSession, setFilterSession] = useState('');
  const [rowsPerPage, setRowsPerPage] = useState<number | string>(10);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    setCurrentPage(1);
  }, [filterMonth, filterYear, filterSession, rowsPerPage]);

  const handleSave = (e) => {
    e.preventDefault();
    if (!formData.link && !formData.notes) return showToast('Please provide either a link or an open-ended question', 'warning');
    
    // Enforce tutor restriction
    if (user.role === 'tutor' && !parseSessions(user.teachingSession).includes(formData.sessionGroup)) {
        return showToast('You can only manage materials for your assigned sessions.', 'error');
    }
    
    if (isEditingId) {
       setDb(p => ({
          ...p,
          // BUGFIX: Tambahkan updatedAt agar LWW tidak kalah ke cloud lama saat edit material
          materials: p.materials.map(m => m.id === isEditingId ? { ...m, ...formData, updatedAt: getSyncTimestamp() } : m)
       }));
       showToast('Material updated successfully');
    } else {
       const newMat = {
         id: generateId('MAT', 'materials'),
         ...formData,
         tutorId: user.id,
         tutorName: user.name,
         date: getTodayDateLocal(),
         submissions: [],
         updatedAt: getSyncTimestamp(), // FIX BUG 3: timestamp untuk LWW
       };
       setDb(p => ({ ...p, materials: [...(p.materials || []), newMat] }));
       showToast('Material posted successfully');
    }
    
    setIsAdding(false);
    setIsEditingId(null);
    setFormData({ title: '', sessionGroup: defaultSession, link: '', notes: '' });
  };

  const editMaterial = (mat) => {
     setFormData({ title: mat.title, sessionGroup: mat.sessionGroup, link: mat.link, notes: mat.notes });
     setIsEditingId(mat.id);
     setIsAdding(true);
     const contentEl = document.querySelector('main'); 
     setTimeout(() => { contentEl?.scrollTo({ top: 0, behavior: 'smooth' }); }, 50);
  };

  const toggleCheck = (matId, subIdx) => {
    setDb(p => {
      const newMats = [...(p.materials || [])];
      const matIdx = newMats.findIndex(m => m.id === matId);
      if (matIdx > -1) {
        newMats[matIdx] = { ...newMats[matIdx] }; // deep copy material
        newMats[matIdx].submissions = [...(newMats[matIdx].submissions || [])]; // deep copy submissions
        newMats[matIdx].submissions[subIdx] = { ...newMats[matIdx].submissions[subIdx] }; // deep copy submission
        newMats[matIdx].submissions[subIdx].checked = !newMats[matIdx].submissions[subIdx].checked;
        // BUGFIX (BUG 8): Update updatedAt di parent material agar LWW tahu versi ini lebih baru dari cloud
        newMats[matIdx].updatedAt = getLocalTimestamp();
      }
      return { ...p, materials: newMats };
    });
  };

  const handleReply = (matId, subIdx) => {
     const text = replyTexts[`${matId}-${subIdx}`]?.trim();
     if (!text) return;
     setDb(p => {
        const newMats = [...(p.materials || [])];
        const matIdx = newMats.findIndex(m => m.id === matId);
        if (matIdx > -1) {
           newMats[matIdx] = { ...newMats[matIdx] }; // deep copy material
           newMats[matIdx].submissions = [...(newMats[matIdx].submissions || [])]; // deep copy submissions
           newMats[matIdx].submissions[subIdx] = { ...newMats[matIdx].submissions[subIdx] }; // deep copy submission
           
           newMats[matIdx].submissions[subIdx].replies = [...(newMats[matIdx].submissions[subIdx].replies || []), {
               senderRole: 'tutor',
               senderName: user.name,
               text,
               date: (() => { const _n = new Date(); return `${String(_n.getDate()).padStart(2,'0')}/${String(_n.getMonth()+1).padStart(2,'0')}/${_n.getFullYear()}, ${String(_n.getHours()).padStart(2,'0')}:${String(_n.getMinutes()).padStart(2,'0')}`; })()
           }];
           newMats[matIdx].submissions[subIdx].checked = true; // Auto-check saat tutor mereply
           newMats[matIdx].updatedAt = getLocalTimestamp(); // BUGFIX (BUG 8): LWW parent material
        }
        return { ...p, materials: newMats };
     });
     setReplyTexts(prev => ({ ...prev, [`${matId}-${subIdx}`]: '' }));
  };

  const myMats = (db.materials || []).filter(m => {
     // FIX #4: Gunakan fuzzy match agar tutor bisa melihat materi yang sama yang terlihat siswa.
     // Mencegah kasus di mana tutor tidak bisa lihat/kelola materi yang sudah terlihat siswa
     // akibat perbedaan minor pada sessionGroup (spasi, nama dari c.name vs SESSIONS[]).
     if (user.role !== 'admin') {
       const _tSessions = parseSessions(user.teachingSession).map(s => s.toLowerCase());
       const _mSession = (m.sessionGroup || '').toLowerCase();
       const _sessionOk = _tSessions.length > 0 && _mSession && _tSessions.some(ts => ts === _mSession || ts.includes(_mSession) || _mSession.includes(ts));
       if (!_sessionOk) return false;
     }
     
     if (filterMonth !== 'All') {
        const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
        if (!m.date.startsWith(prefix)) return false;
     } else {
        if (!m.date.startsWith(String(filterYear))) return false;
     }

     if (filterSession && m.sessionGroup !== filterSession) return false;
     
     return true;
  }).reverse();

  // Pagination Logic
  const isAll = rowsPerPage === 'All';
  const rowsNum = isAll ? myMats.length : Number(rowsPerPage);
  const totalPages = isAll ? 1 : Math.ceil(myMats.length / (rowsNum || 1));
  const paginatedData = isAll ? myMats : myMats.slice((currentPage - 1) * rowsNum, currentPage * rowsNum);

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white mb-1">Materials & Tasks</h2><p className="text-gray-400 text-sm">Upload and manage learning materials and assignments.</p></div>
        <Button onClick={() => { setIsAdding(!isAdding); setIsEditingId(null); setFormData({ title: '', sessionGroup: defaultSession, link: '', notes: '' }); }} icon={isAdding ? X : Plus} variant={isAdding ? 'secondary' : 'primary'}>{isAdding ? 'Cancel' : 'Add Material'}</Button>
      </div>

      {isAdding && (
        <Card>
          <h3 className="text-lg font-bold text-white mb-4">{isEditingId ? 'Edit Material' : 'Publish Material'}</h3>
          <form onSubmit={handleSave} className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <Input label="Material Title" value={formData.title} onChange={v => setFormData({...formData, title: v})} required />
            <Input 
               label="Session Group" 
               type="select" 
               options={user.role === 'tutor' ? parseSessions(user.teachingSession) : SESSIONS} 
               value={formData.sessionGroup} 
               onChange={v => setFormData({...formData, sessionGroup: v})} 
               required 
            />
            <div className="md:col-span-2">
               <Input label="Link (YouTube / Website) - Optional" type="url" value={formData.link} onChange={v => setFormData({...formData, link: v})} placeholder="https://..." />
            </div>
            <div className="md:col-span-2">
              <label className="block text-sm text-gray-400 mb-1.5">Open-ended Question / Instructions</label>
              <textarea className="w-full bg-[#0B0F19] border border-gray-700 rounded-lg p-3 text-white h-24 focus:border-[#00D4FF]" value={formData.notes} onChange={e => setFormData({...formData, notes: e.target.value})} placeholder="e.g. Watch the video, or answer this specific question..." />
            </div>
            <div className="md:col-span-2 flex justify-end gap-2">
              <Button variant="ghost" onClick={() => { setIsAdding(false); setIsEditingId(null); }}>Cancel</Button>
              <Button type="submit">{isEditingId ? 'Save Changes' : 'Publish Material'}</Button>
            </div>
          </form>
        </Card>
      )}

      <Card className="p-0 flex flex-col">
         {/* Filter Row */}
         <div className="p-4 bg-[#0A0E17] border-b border-gray-800 flex flex-col md:flex-row justify-between items-center gap-3">
            <div className="flex gap-2 w-full md:w-auto">
               <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
                  <option value="All">All Months</option>
                  {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
               </select>
               <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
            </div>
            <select 
               className="w-full md:w-48 bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-[#00D4FF]" 
               value={filterSession} 
               onChange={(e) => setFilterSession(e.target.value)}
            >
               {user.role === 'tutor' ? (
                 <>
                   <option value="">All My Sessions</option>
                   {parseSessions(user.teachingSession).map(s => <option key={s} value={s}>{s}</option>)}
                 </>
               ) : (
                 <>
                   <option value="">All Sessions</option>
                   {SESSIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                 </>
               )}
            </select>
         </div>
      </Card>

      <div className="space-y-4">
        {paginatedData.map(mat => {
          const preview = getLinkPreview(mat.link);
          return (
            <Card key={mat.id} className="p-0 overflow-hidden border border-gray-800 shadow-xl">
               <div className="flex flex-col md:flex-row border-b border-gray-800 bg-[#151B26]">
                  {mat.link && (
                     <div className="w-full md:w-64 h-48 md:h-auto bg-black flex-shrink-0 relative group">
                        {preview.type === 'youtube' ? (
                           playingVideos[mat.id] ? (
                              <iframe 
                                 src={`https://www.youtube.com/embed/${preview.videoId}?autoplay=1`} 
                                 title="YouTube video player" 
                                 frameBorder="0" 
                                 allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" 
                                 allowFullScreen
                                 className="absolute inset-0 w-full h-full"
                              ></iframe>
                           ) : (
                              <>
                                 <img src={preview.src} alt="Thumbnail" className="w-full h-full object-cover opacity-80 group-hover:opacity-100 transition-opacity" />
                                 <button type="button" onClick={() => setPlayingVideos(p => ({...p, [mat.id]: true}))} className="absolute inset-0 flex items-center justify-center w-full h-full focus:outline-none">
                                    <div className="w-12 h-12 bg-red-600 rounded-full flex items-center justify-center shadow-[0_0_20px_rgba(220,38,38,0.5)] group-hover:scale-110 transition-transform">
                                       <div className="w-0 h-0 border-t-8 border-b-8 border-l-[14px] border-transparent border-l-white ml-1"></div>
                                    </div>
                                 </button>
                              </>
                           )
                        ) : preview.type === 'website' ? (
                           <>
                              <img src={preview.src} alt="Website Thumbnail" onError={(e) => { e.currentTarget.style.display = 'none'; }} className="w-full h-full object-cover opacity-60 group-hover:opacity-90 transition-opacity object-top" />
                              <a href={mat.link} target="_blank" rel="noopener noreferrer" className="absolute inset-0 flex flex-col items-center justify-center bg-black/30 group-hover:bg-black/10 transition-colors">
                                  <div className="bg-[#00D4FF] text-[#0B0F19] px-4 py-2 rounded-xl font-bold flex items-center gap-2 shadow-[0_0_15px_rgba(0,212,255,0.4)] group-hover:scale-105 transition-transform">
                                     <LinkIcon size={16} /> Open Link
                                  </div>
                              </a>
                           </>
                        ) : (
                           <a href={mat.link} target="_blank" rel="noopener noreferrer" className="w-full h-full flex flex-col items-center justify-center text-blue-400 hover:text-blue-300 hover:bg-blue-900/20 transition-colors gap-3 py-10">
                              <LinkIcon size={40} />
                              <span className="text-sm font-semibold">Open External Link</span>
                           </a>
                        )}
                     </div>
                  )}
                  
                  <div className="p-5 flex-1 flex flex-col">
                     <div className="flex justify-between items-start mb-4 border-b border-gray-800/50 pb-3">
                         <div>
                             <h3 className="text-xl font-bold text-white mb-1 leading-tight">{mat.title}</h3>
                             <div className="flex items-center gap-3 text-sm text-gray-400">
                                <span className="text-[#00D4FF] font-medium">{mat.sessionGroup}</span>
                                <span>•</span>
                                <span>{mat.date} by {mat.tutorName}</span>
                             </div>
                         </div>
                         <div className="flex gap-2">
                             <button onClick={() => editMaterial(mat)} className="text-blue-400 p-2.5 hover:bg-blue-500/10 rounded-lg transition-colors" title="Edit Material"><Edit2 size={18} /></button>
                             <button onClick={() => softDelete('materials', mat.id, 'Material')} className="text-red-400 p-2.5 hover:bg-red-500/10 rounded-lg transition-colors" title="Delete Material"><Trash2 size={18} /></button>
                         </div>
                     </div>
                     <p className="text-gray-300 text-sm whitespace-pre-wrap leading-relaxed">{mat.notes}</p>
                  </div>
               </div>

               <div className="p-5 bg-[#0B0F19]">
                  <h4 className="text-sm font-bold text-white mb-3 flex items-center gap-2"><MessageCircle size={16} className="text-[#00D4FF]" /> Student Submissions ({(mat.submissions || []).length})</h4>
                  <div className="space-y-4">
                     {(mat.submissions || []).map((sub, idx) => (
                        <div key={idx} className={`rounded-xl border flex flex-col overflow-hidden transition-colors ${sub.checked ? 'border-emerald-500/30' : 'border-gray-700'}`}>
                           
                           <div className={`p-3 flex justify-between items-center ${sub.checked ? 'bg-emerald-500/10' : 'bg-[#1A2234]'}`}>
                              <p className="font-bold text-white text-sm">
                                {sub.studentName} <span className="text-xs text-gray-500 font-normal ml-2">{sub.date}</span>
                              </p>
                              <button 
                                onClick={() => toggleCheck(mat.id, idx)} 
                                className={`p-1.5 rounded-lg transition-colors flex-shrink-0 shadow-sm ${sub.checked ? 'bg-emerald-500 text-white shadow-emerald-500/20' : 'bg-gray-700 text-gray-400 hover:text-white hover:bg-gray-600'}`}
                                title={sub.checked ? 'Mark as Unchecked' : 'Mark as Checked'}
                              >
                                 <Check size={14} />
                              </button>
                           </div>
                           
                           <div className="p-4 space-y-3 bg-[#0B0F19]">
                              <div className="flex flex-col items-start">
                                 <div className="bg-[#151B26] border border-gray-800 p-3 rounded-lg rounded-tl-none inline-block max-w-[85%]">
                                    <span className="text-xs font-bold text-blue-400 block mb-1">{sub.studentName}</span>
                                    <p className="text-sm text-gray-300 whitespace-pre-wrap">{sub.text}</p>
                                 </div>
                              </div>
                              {(sub.replies || []).map((r, i) => (
                                 <div key={i} className={`flex flex-col ${r.senderRole === 'tutor' ? 'items-end' : 'items-start'}`}>
                                    <div className={`p-3 rounded-lg inline-block max-w-[85%] ${r.senderRole === 'tutor' ? 'bg-[#00D4FF]/10 border border-[#00D4FF]/20 rounded-tr-none' : 'bg-[#151B26] border border-gray-800 rounded-tl-none'}`}>
                                       <span className={`text-[11px] font-bold block mb-1 ${r.senderRole === 'tutor' ? 'text-[#00D4FF] text-right' : 'text-blue-400'}`}>
                                         {r.senderRole === 'tutor' ? 'Me' : r.senderName} <span className="text-gray-500 font-normal ml-1">{r.date}</span>
                                       </span>
                                       <p className="text-sm text-gray-300 whitespace-pre-wrap">{r.text}</p>
                                    </div>
                                 </div>
                              ))}
                           </div>
                           
                           <div className="p-3 bg-[#151B26] border-t border-gray-800 flex gap-2">
                              <input 
                                 type="text" 
                                 className="flex-1 bg-[#0B0F19] border border-gray-700 rounded-lg px-4 py-2 text-white text-sm focus:border-[#00D4FF] focus:outline-none"
                                 placeholder="Write a feedback or reply..."
                                 value={replyTexts[`${mat.id}-${idx}`] || ''}
                                 onChange={e => setReplyTexts(p => ({...p, [`${mat.id}-${idx}`]: e.target.value}))}
                                 onKeyDown={e => { if (e.key === 'Enter') handleReply(mat.id, idx); }}
                              />
                              <Button onClick={() => handleReply(mat.id, idx)} className="px-4 text-sm" icon={MessageSquare}>Reply</Button>
                           </div>
                        </div>
                     ))}
                     {(mat.submissions || []).length === 0 && <p className="text-xs text-gray-500 italic px-2">No submissions yet.</p>}
                  </div>
               </div>
            </Card>
          );
        })}
        {paginatedData.length === 0 && <div className="p-8 text-center text-gray-500 bg-[#151B26] rounded-xl border border-gray-800">No materials posted matching your filters.</div>}
      </div>

      {/* Pagination Footer */}
      {myMats.length > 0 && (
          <div className="p-4 bg-[#0A0E17] border border-gray-800 rounded-xl flex flex-col sm:flex-row justify-between items-center gap-4 text-sm text-gray-400 mt-4">
             <div className="flex items-center gap-2">
               <span>Show</span>
               <select value={rowsPerPage} onChange={(e) => setRowsPerPage(e.target.value === 'All' ? 'All' : Number(e.target.value))} className="bg-[#151B26] border border-gray-700 rounded px-2 py-1 text-white focus:outline-none focus:border-[#00D4FF] cursor-pointer">
                 <option value={10}>10</option>
                 <option value={20}>20</option>
                 <option value={50}>50</option>
                 <option value="All">All</option>
               </select>
               <span>entries (Total: {myMats.length})</span>
             </div>
             
             {!isAll && totalPages > 1 && (
               <div className="flex items-center gap-2">
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === 1} onClick={() => setCurrentPage(p => p - 1)}>Prev</Button>
                 <span className="px-3 py-1.5 text-white font-medium">{currentPage} / {totalPages}</span>
                 <Button variant="ghost" className="px-3 py-1.5 h-auto text-xs bg-[#151B26] border border-gray-700 hover:bg-gray-800" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => p + 1)}>Next</Button>
               </div>
             )}
          </div>
      )}
    </div>
  );
}

function StudentMaterialsModule({ db, setDb, user, showToast, language = 'en' }) {
  const [comment, setComment] = useState({});
  const [filterMonth, setFilterMonth] = useState<number | string>(new Date().getMonth() + 1);
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  const [playingVideos, setPlayingVideos] = useState({});
  const student = db.students.find(s => s.id === user.studentId);
  const myGroup = student ? getStudentSession(student) : '';
  // FIX F1: sessionGroup fuzzy match (sama seperti di journals)
  const sessionMatches = (mGroup, sGroup) => {
    if (!mGroup || !sGroup) return false;
    if (mGroup === sGroup) return true;
    const a = mGroup.toLowerCase(), b = sGroup.toLowerCase();
    return a.includes(b) || b.includes(a);
  };
  const myMats = (db.materials || []).filter(m => {
     if (!sessionMatches(m.sessionGroup, myGroup)) return false;
     if (filterMonth !== 'All') {
        const prefix = `${filterYear}-${String(filterMonth).padStart(2, '0')}`;
        return m.date.startsWith(prefix);
     }
     return m.date.startsWith(String(filterYear));
  }).reverse();

  const handleSubmitComment = (matId, existingSubIdx = -1) => {
    const text = comment[matId]?.trim();
    if (!text) return showToast('Please write your comment or submission first', 'warning');
    // FIX #5: Guard jika data siswa belum dimuat dari cloud — jangan simpan submission
    // dengan studentId undefined karena akan menyebabkan duplikasi submission tak terbatas.
    if (!student?.id) return showToast(language === 'id' ? 'Data siswa belum dimuat, coba lagi.' : 'Student data not loaded yet, please try again.', 'warning');
    
    setDb(p => {
      const newMats = [...(p.materials || [])];
      const matIdx = newMats.findIndex(m => m.id === matId);
      if (matIdx > -1) {
        newMats[matIdx] = { ...newMats[matIdx] }; // deep copy material
        newMats[matIdx].submissions = [...(newMats[matIdx].submissions || [])]; // deep copy submissions
        
        if (existingSubIdx > -1) {
           // Append reply
           newMats[matIdx].submissions[existingSubIdx] = { ...newMats[matIdx].submissions[existingSubIdx] }; // deep copy submission
           newMats[matIdx].submissions[existingSubIdx].replies = [...(newMats[matIdx].submissions[existingSubIdx].replies || []), {
               senderRole: 'student',
               senderName: student?.name || 'Student',
               text,
               date: (() => { const _n = new Date(); return `${String(_n.getDate()).padStart(2,'0')}/${String(_n.getMonth()+1).padStart(2,'0')}/${_n.getFullYear()}, ${String(_n.getHours()).padStart(2,'0')}:${String(_n.getMinutes()).padStart(2,'0')}`; })()
           }];
           newMats[matIdx].submissions[existingSubIdx].checked = false; // Mark as unread/unchecked for tutor
        } else {
           // Initial Submission
           newMats[matIdx].submissions.push({
             studentId: student?.id,
             studentName: student?.name || 'Student',
             text,
             date: (() => { const _n = new Date(); return `${String(_n.getDate()).padStart(2,'0')}/${String(_n.getMonth()+1).padStart(2,'0')}/${_n.getFullYear()}, ${String(_n.getHours()).padStart(2,'0')}:${String(_n.getMinutes()).padStart(2,'0')}`; })(),
             checked: false,
             replies: []
           });
        }
        // BUGFIX (BUG 8): Update updatedAt parent material agar LWW tahu ada submission/reply baru
        newMats[matIdx].updatedAt = getLocalTimestamp();
      }
      return { ...p, materials: newMats };
    });
    
    showToast('Submission sent successfully!');
    setComment(p => ({ ...p, [matId]: '' }));
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
        <div>
          <h2 className="text-2xl font-bold text-white mb-1">{language === 'id' ? 'Materi & Tugas Saya' : 'My Materials & Tasks'}</h2>
          <p className="text-sm text-gray-400">{language === 'id' ? 'Lihat materi, tonton video, dan kirim tugas/komentar Anda di bawah ini.' : 'View materials, watch videos, and submit your tasks/comments below.'}</p>
        </div>
        <div className="flex gap-2 w-full sm:w-auto">
            <select className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white focus:border-[#00D4FF] flex-1 sm:flex-none" value={filterMonth} onChange={e => setFilterMonth(e.target.value === 'All' ? 'All' : Number(e.target.value))}>
              <option value="All">{language === 'id' ? 'Semua Bulan' : 'All Months'}</option>
              {MONTHS.map((m, i) => <option key={i} value={i + 1}>{m}</option>)}
            </select>
            <input type="number" className="bg-[#151B26] border border-gray-700 rounded-lg px-3 py-2 text-sm text-white w-24 focus:border-[#00D4FF]" value={filterYear} onChange={e => setFilterYear(Number(e.target.value))} />
        </div>
      </div>

      <div className="space-y-6">
        {myMats.map((mat, matIdx) => {
          const preview = getLinkPreview(mat.link);
          const mySubIdx = (mat.submissions || []).findIndex(s => s.studentId === student?.id);
          const mySub = mySubIdx > -1 ? mat.submissions[mySubIdx] : null;
          
          return (
            <Card key={mat.id} className="p-0 overflow-hidden border border-gray-800 shadow-xl">
               <div className="flex flex-col md:flex-row border-b border-gray-800 bg-[#151B26]">
                  {mat.link && (
                     <div className="w-full md:w-64 h-48 md:h-auto bg-black flex-shrink-0 relative group">
                        {preview.type === 'youtube' ? (
                           playingVideos[mat.id] ? (
                              <iframe 
                                 src={`https://www.youtube.com/embed/${preview.videoId}?autoplay=1`} 
                                 title="YouTube video player" 
                                 frameBorder="0" 
                                 allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" 
                                 allowFullScreen
                                 className="absolute inset-0 w-full h-full"
                              ></iframe>
                           ) : (
                              <>
                                 <img src={preview.src} alt="Thumbnail" className="w-full h-full object-cover opacity-80 group-hover:opacity-100 transition-opacity" />
                                 <button type="button" onClick={() => setPlayingVideos(p => ({...p, [mat.id]: true}))} className="absolute inset-0 flex items-center justify-center w-full h-full focus:outline-none">
                                    <div className="w-12 h-12 bg-red-600 rounded-full flex items-center justify-center shadow-[0_0_20px_rgba(220,38,38,0.5)] group-hover:scale-110 transition-transform">
                                       <div className="w-0 h-0 border-t-8 border-b-8 border-l-[14px] border-transparent border-l-white ml-1"></div>
                                    </div>
                                 </button>
                              </>
                           )
                        ) : preview.type === 'website' ? (
                           <>
                              <img src={preview.src} alt="Website Thumbnail" onError={(e) => { e.currentTarget.style.display = 'none'; }} className="w-full h-full object-cover opacity-60 group-hover:opacity-90 transition-opacity object-top" />
                              <a href={mat.link} target="_blank" rel="noopener noreferrer" className="absolute inset-0 flex flex-col items-center justify-center bg-black/30 group-hover:bg-black/10 transition-colors">
                                  <div className="bg-[#00D4FF] text-[#0B0F19] px-4 py-2 rounded-xl font-bold flex items-center gap-2 shadow-[0_0_15px_rgba(0,212,255,0.4)] group-hover:scale-105 transition-transform">
                                     <LinkIcon size={16} /> {language === 'id' ? 'Buka Tautan' : 'Open Link'}
                                  </div>
                              </a>
                           </>
                        ) : (
                           <a href={mat.link} target="_blank" rel="noopener noreferrer" className="w-full h-full flex flex-col items-center justify-center text-blue-400 hover:text-blue-300 hover:bg-blue-900/20 transition-colors gap-3 py-10">
                              <LinkIcon size={40} />
                              <span className="text-sm font-semibold">{language === 'id' ? 'Buka Tautan Eksternal' : 'Open External Link'}</span>
                           </a>
                        )}
                     </div>
                  )}
                  
                  <div className="p-5 flex-1">
                     <div className="mb-4 border-b border-gray-800/50 pb-3">
                        <div className="flex items-center gap-2 mb-1">
                           <span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-[#00D4FF]/10 border border-[#00D4FF]/20 text-[11px] font-black text-[#00D4FF] shrink-0">{String(matIdx + 1).padStart(2, '0')}</span>
                           <h3 className="text-xl font-bold text-white leading-tight">{mat.title}</h3>
                        </div>
                        <p className="text-xs text-gray-400">{language === 'id' ? 'Diposting pada' : 'Posted on'} {mat.date} {language === 'id' ? 'oleh' : 'by'} <span className="text-[#00D4FF]">{mat.tutorName}</span></p>
                     </div>
                     <p className="text-gray-300 text-sm whitespace-pre-wrap leading-relaxed">{mat.notes}</p>
                  </div>
               </div>

               <div className="p-5 bg-[#0B0F19]">
                  <h4 className="text-sm font-bold text-white mb-3 flex items-center gap-2">
                     <MessageCircle size={16} className="text-[#00D4FF]" /> 
                     {mySub ? (language === 'id' ? 'Pengumpulan Tugas & Umpan Balik' : 'Task Submission & Feedback') : (language === 'id' ? 'Kumpulkan Tugas Anda' : 'Submit Your Task')}
                  </h4>
                  
                  {mySub ? (
                     <div className="border border-gray-800 rounded-xl overflow-hidden bg-[#151B26]">
                        <div className="p-3 bg-[#1A2234] border-b border-gray-800 flex justify-between items-center">
                           <span className="text-xs font-bold text-white">{language === 'id' ? 'Utas Percakapan' : 'Conversation Thread'}</span>
                           {mySub.checked ? (
                              <span className="text-[11px] bg-emerald-500/20 text-emerald-400 px-2 py-1 rounded-full font-bold flex items-center gap-1 shadow-sm"><Check size={12}/> {language === 'id' ? 'Diperiksa oleh Tutor' : 'Checked by Tutor'}</span>
                           ) : (
                              <span className="text-[11px] bg-amber-500/20 text-amber-400 px-2 py-1 rounded-full font-bold">{language === 'id' ? 'Menunggu Ulasan' : 'Pending Review'}</span>
                           )}
                        </div>
                        <div className="p-4 space-y-3 max-h-[300px] overflow-y-auto custom-scrollbar">
                           <div className="flex flex-col items-end">
                              <div className="bg-[#00D4FF]/10 border border-[#00D4FF]/20 p-3 rounded-lg rounded-tr-none inline-block max-w-[85%]">
                                 <span className="text-[11px] font-bold text-[#00D4FF] block mb-1 text-right">{language === 'id' ? 'Saya' : 'Me'} <span className="text-gray-500 font-normal ml-1">{mySub.date}</span></span>
                                 <p className="text-sm text-gray-200 whitespace-pre-wrap">{mySub.text}</p>
                              </div>
                           </div>
                           {(mySub.replies || []).map((r, i) => (
                              <div key={i} className={`flex flex-col ${r.senderRole === 'student' ? 'items-end' : 'items-start'}`}>
                                 <div className={`p-3 rounded-lg inline-block max-w-[85%] ${r.senderRole === 'student' ? 'bg-[#00D4FF]/10 border border-[#00D4FF]/20 rounded-tr-none' : 'bg-[#1A2234] border border-gray-700 rounded-tl-none'}`}>
                                    <span className={`text-[11px] font-bold block mb-1 ${r.senderRole === 'student' ? 'text-[#00D4FF] text-right' : 'text-purple-400'}`}>
                                       {r.senderRole === 'student' ? (language === 'id' ? 'Saya' : 'Me') : r.senderName} <span className="text-gray-500 font-normal ml-1">{r.date}</span>
                                    </span>
                                    <p className="text-sm text-gray-200 whitespace-pre-wrap">{r.text}</p>
                                 </div>
                              </div>
                           ))}
                        </div>
                        
                        <div className="p-3 bg-[#0B0F19] border-t border-gray-800 flex gap-2">
                           <input 
                              type="text" 
                              value={comment[mat.id] || ''} 
                              onChange={e => setComment(p => ({...p, [mat.id]: e.target.value}))}
                              placeholder={language === 'id' ? "Balas ke tutor..." : "Reply to tutor..."}
                              className="flex-1 bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2 text-white text-sm focus:outline-none focus:border-[#00D4FF]"
                              onKeyDown={e => { if (e.key === 'Enter') handleSubmitComment(mat.id, mySubIdx); }}
                           />
                           <Button onClick={() => handleSubmitComment(mat.id, mySubIdx)} className="px-4 text-sm" icon={MessageSquare}>{language === 'id' ? 'Balas' : 'Reply'}</Button>
                        </div>
                     </div>
                  ) : (
                     <div className="flex gap-3 mt-2">
                        <input 
                           type="text" 
                           value={comment[mat.id] || ''} 
                           onChange={e => setComment(p => ({...p, [mat.id]: e.target.value}))}
                           placeholder={language === 'id' ? "Tulis jawaban atau tugas Anda di sini..." : "Write your answer or task here..."}
                           className="flex-1 bg-[#151B26] border border-gray-700 rounded-lg px-4 py-2.5 text-white text-sm focus:outline-none focus:border-[#00D4FF]"
                           onKeyDown={e => { if (e.key === 'Enter') handleSubmitComment(mat.id, -1); }}
                        />
                        <Button onClick={() => handleSubmitComment(mat.id, -1)} className="px-6 shadow-md">{language === 'id' ? 'Kirim Tugas' : 'Submit Task'}</Button>
                     </div>
                  )}
               </div>
            </Card>
          );
        })}
        {myMats.length === 0 && (
           <EmptyState
              icon={BookOpen}
              title={language === 'id' ? 'Belum Ada Materi' : 'No Materials Yet'}
              description={language === 'id' ? 'Tutor Anda belum memposting materi atau tugas untuk sesi ini.' : "Your tutor hasn't posted any materials or tasks for your session."}
              className="py-12 bg-[#151B26] border border-gray-800 rounded-xl shadow-md"
           />
        )}
      </div>
    </div>
  );
}

if (typeof document !== 'undefined' && !document.getElementById('ecg-styles')) {
  const style = document.createElement('style');
  style.id = 'ecg-styles';
  style.textContent = `
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    
    .font-sans { font-family: 'Inter', sans-serif; }
    
    .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
    .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background: #334155; border-radius: 10px; }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #475569; }
    
    @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
    .animation-fade-in { animation: fadeIn 0.3s ease-out forwards; }
    
    @keyframes scaleIn { from { opacity: 0; transform: scale(0.5); } to { opacity: 1; transform: scale(1); } }
    .receipt-pop { animation: scaleIn 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards; }
    
    @media print {
      @page {
        size: A4 portrait;
        margin: 5mm; 
      }
      body { 
        background: white !important; 
        -webkit-print-color-adjust: exact !important; 
        print-color-adjust: exact !important; 
      }
      * { 
        text-shadow: none !important; 
        box-shadow: none !important; 
        letter-spacing: normal !important; /* FIX #2: BROKEN LETTER SPACING */
      }
      
      /* Smart PDF Compression Rules to Force 1 Page (FIX #1 & #4) */
      #report-print, #slip-print {
        width: 100%;
        zoom: 0.92; /* Safer scaling than transform for print layouts */
      }
      
      /* Aggressively reduce gaps for print to fit on 1 page */
      .mb-8 { margin-bottom: 0.5rem !important; }
      .mb-6 { margin-bottom: 0.5rem !important; }
      .mt-8 { margin-top: 0.5rem !important; }
      .mt-10 { margin-top: 0.5rem !important; }
      .p-8 { padding: 1rem !important; }
      .pb-4 { padding-bottom: 0.25rem !important; }
      .pt-4 { padding-top: 0.25rem !important; }
      .pt-6 { padding-top: 0.25rem !important; }
      .gap-4 { gap: 0.25rem !important; }
      td, th { padding: 0.2rem 0.3rem !important; }
      
      .signature-section { 
        page-break-inside: avoid !important; 
        break-inside: avoid !important;
        margin-top: 0 !important;
      }
      
      .print-hidden, aside, header.sticky { display: none !important; }
      .print-border { border: 1px solid #e5e7eb !important; }
      .print-border-highlight { border: 2px solid #bae6fd !important; }
      .print-border-gray-200 { border-color: #e5e7eb !important; }
      .print-border-gray-300 { border-color: #d1d5db !important; }
      .print-text-gray-900 { color: #111827 !important; }
      .print-text-black { color: #000 !important; }
      .print-bg-white { background-color: #fff !important; }
      .print-shadow-none { box-shadow: none !important; }
      
      html, body, #root, .h-screen, .overflow-y-auto, .flex-1, main {
        height: auto !important;
        overflow: visible !important;
        position: static !important;
      }
      
      tr, .break-inside-avoid, .signature-section { page-break-inside: avoid !important; break-inside: avoid !important; }
      h2, h3, h4 { page-break-after: avoid !important; }
      
      table.print-table { border-collapse: collapse; width: 100%; }
      thead.table-header-group { display: table-header-group; }
      tfoot.table-footer-group { display: table-footer-group; }
      
      /* FIX #3: Removed print-page-number::after CSS to prevent double page numbers */
    }
  `;
  document.head.appendChild(style);
}