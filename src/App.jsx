import { useState, useEffect, useRef } from 'react'
import { db } from './firebase'
import { collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc, query, where, getDocs, serverTimestamp } from 'firebase/firestore'
import * as XLSX from 'xlsx'

const PRIORITIES = {
  high: { label: 'גבוהה', color: '#fef2f2', text: '#dc2626', dot: '#ef4444', order: 3 },
  medium: { label: 'בינונית', color: '#fffbeb', text: '#d97706', dot: '#f59e0b', order: 2 },
  low: { label: 'נמוכה', color: '#f0fdf4', text: '#16a34a', dot: '#22c55e', order: 1 }
}
const STATUSES = {
  todo: { label: 'לביצוע', color: '#f8fafc', text: '#475569', dot: '#94a3b8' },
  inprogress: { label: 'בביצוע', color: '#eff6ff', text: '#2563eb', dot: '#3b82f6' },
  done: { label: 'הושלם', color: '#f0fdf4', text: '#16a34a', dot: '#22c55e' }
}
const today = new Date().toISOString().split('T')[0]
const norm = s => s?.trim().toLowerCase() || ''
const avatarColors = ['#0ea5e9','#10b981','#f59e0b','#06b6d4','#ec4899','#14b8a6','#84cc16','#3b82f6']
const getColor = name => avatarColors[(name?.charCodeAt(0) || 0) % avatarColors.length]
const DEFAULT_COL_WIDTHS = { project: 130, title: 220, assignee: 150, priority: 120, status: 140, dueDate: 130, notes: 200, actions: 100 }

const Avatar = ({ name, size = 28 }) => (
  <span style={{ width: size, height: size, borderRadius: '50%', background: getColor(name), color: 'white', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: size * 0.4, fontWeight: 700, flexShrink: 0 }}>
    {name?.[0]?.toUpperCase()}
  </span>
)

const Badge = ({ label, color, text, dot }) => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: color, color: text, padding: '4px 12px', borderRadius: 999, fontSize: 12, fontWeight: 700, border: `1px solid ${dot}40` }}>
    <span style={{ width: 7, height: 7, borderRadius: '50%', background: dot }} />{label}
  </span>
)

const Tooltip = ({ text }) => {
  const [show, setShow] = useState(false)
  const [pos, setPos] = useState({ x: 0, y: 0 })
  if (!text) return <span style={{ color: '#cbd5e1' }}>—</span>
  return (
    <span onMouseEnter={e => { setShow(true); setPos({ x: e.clientX, y: e.clientY }) }} onMouseLeave={() => setShow(false)} style={{ cursor: 'default' }}>
      <span style={{ color: '#0f172a', fontSize: 14, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{text}</span>
      {show && <div style={{ position: 'fixed', top: pos.y + 10, left: pos.x, background: '#1e293b', color: '#f1f5f9', padding: '10px 14px', borderRadius: 10, fontSize: 13, maxWidth: 320, zIndex: 9999, boxShadow: '0 8px 30px rgba(0,0,0,0.25)', lineHeight: 1.6, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{text}</div>}
    </span>
  )
}

const StatCard = ({ label, value, color, bg, icon }) => (
  <div style={{ background: 'white', borderRadius: 14, padding: '1rem 1.2rem', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 12, boxShadow: '0 1px 3px rgba(0,0,0,0.04)' }}>
    <div style={{ width: 44, height: 44, borderRadius: 12, background: bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>{icon}</div>
    <div>
      <div style={{ fontSize: 24, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
      <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 3 }}>{label}</div>
    </div>
  </div>
)

export default function App() {
  const [currentUser, setCurrentUser] = useState(null)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [tasks, setTasks] = useState([])
  const [deletedTasks, setDeletedTasks] = useState([])
  const [users, setUsers] = useState([])
  const [selectedUser, setSelectedUser] = useState(null)
  const [showForm, setShowForm] = useState(false)
  const [editTask, setEditTask] = useState(null)
  const [filterStatus, setFilterStatus] = useState('all')
  const [filterPriority, setFilterPriority] = useState('all')
  const [filterProject, setFilterProject] = useState('all')
  const [search, setSearch] = useState('')
  const [sortField, setSortField] = useState('')
  const [sortDir, setSortDir] = useState('asc')
  const [notifications, setNotifications] = useState([])
  const [showNotifications, setShowNotifications] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [selectedUserForPassword, setSelectedUserForPassword] = useState(null)
  const [newPassword, setNewPassword] = useState('')
  const [showDeleted, setShowDeleted] = useState(false)
  const [showDashboard, setShowDashboard] = useState(false)
  const [colWidths, setColWidths] = useState(() => {
    try { const s = localStorage.getItem('colWidths'); return s ? JSON.parse(s) : DEFAULT_COL_WIDTHS } catch { return DEFAULT_COL_WIDTHS }
  })
  const [resizing, setResizing] = useState(null)
  const [form, setForm] = useState({ project: '', title: '', assignee: '', assigneeEmail: '', priority: 'medium', status: 'todo', dueDate: '', notes: '' })
  const prevTasksRef = useRef([])
  const fileInputRef = useRef(null)
  const tableRef = useRef(null)

  useEffect(() => {
    const saved = localStorage.getItem('currentUser')
    if (saved) setCurrentUser(JSON.parse(saved))
  }, [])

  useEffect(() => { localStorage.setItem('colWidths', JSON.stringify(colWidths)) }, [colWidths])

  useEffect(() => {
    if (!currentUser) return
    return onSnapshot(collection(db, 'users'), snap => setUsers(snap.docs.map(d => ({ id: d.id, ...d.data() }))))
  }, [currentUser])

  useEffect(() => {
    if (!currentUser) return
    const q = currentUser.role === 'admin' ? collection(db, 'tasks') : query(collection(db, 'tasks'), where('assigneeEmail', '==', norm(currentUser.email)))
    return onSnapshot(q, snap => {
      const newTasks = snap.docs.map(d => ({ id: d.id, ...d.data() }))
      if (prevTasksRef.current.length > 0) {
        if (currentUser.role === 'admin') {
          newTasks.forEach(task => {
            const prev = prevTasksRef.current.find(t => t.id === task.id)
            if (prev && prev.status !== task.status)
              saveNotif(norm(currentUser.email), `${task.assignee || 'עובד'} עדכן "${task.title}" → ${STATUSES[task.status]?.label}`, 'status')
          })
        } else {
          newTasks.forEach(task => {
            const prev = prevTasksRef.current.find(t => t.id === task.id)
            if (!prev) saveNotif(norm(currentUser.email), `משימה חדשה: "${task.title}"`, 'new')
            else if (prev.status !== task.status) saveNotif(norm(currentUser.email), `סטטוס עודכן: "${task.title}" → ${STATUSES[task.status]?.label}`, 'status')
            else if (prev.priority !== task.priority) saveNotif(norm(currentUser.email), `עדיפות שונתה: "${task.title}"`, 'priority')
            else if (prev.dueDate !== task.dueDate) saveNotif(norm(currentUser.email), `תאריך יעד שונה: "${task.title}"`, 'date')
          })
        }
      }
      prevTasksRef.current = newTasks
      setTasks(newTasks)
    })
  }, [currentUser])

  useEffect(() => {
    if (!currentUser || currentUser.role !== 'admin') return
    return onSnapshot(collection(db, 'deleted_tasks'), snap => {
      const now = Date.now()
      setDeletedTasks(snap.docs.map(d => ({ id: d.id, ...d.data() })).filter(t => t.placeholder !== 'true' && t.deletedAt && (now - t.deletedAt) < 24 * 60 * 60 * 1000))
    })
  }, [currentUser])

  useEffect(() => {
    if (!currentUser) return
    return onSnapshot(query(collection(db, 'notifications'), where('userEmail', '==', norm(currentUser.email))), snap => {
      setNotifications(snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0)))
    })
  }, [currentUser])

  const saveNotif = async (userEmail, message, type) => {
    if (!userEmail) return
    await addDoc(collection(db, 'notifications'), { userEmail: norm(userEmail), message, type, read: false, createdAt: serverTimestamp() })
  }

  const login = async () => {
    setError('')
    const snap = await getDocs(collection(db, 'users'))
    const userData = snap.docs.map(d => ({ id: d.id, ...d.data() })).find(u => norm(u.email) === norm(email))
    if (!userData) return setError('אימייל לא קיים במערכת')
    if (userData.password !== password) return setError('סיסמה שגויה')
    setCurrentUser(userData)
    localStorage.setItem('currentUser', JSON.stringify(userData))
  }

  const logout = () => { setCurrentUser(null); localStorage.removeItem('currentUser'); setTasks([]); setNotifications([]); setSelectedUser(null) }

  const toggleNotifications = () => setShowNotifications(s => !s)

  const markAllRead = async () => {
    for (const n of notifications.filter(n => !n.read)) await updateDoc(doc(db, 'notifications', n.id), { read: true })
  }

  const markOneRead = async n => {
    if (!n.read) await updateDoc(doc(db, 'notifications', n.id), { read: true })
  }

  const clearAllNotifications = async () => { for (const n of notifications) await deleteDoc(doc(db, 'notifications', n.id)) }

  const saveTask = async () => {
    if (!form.title.trim()) return
    const taskData = { ...form, assigneeEmail: norm(form.assigneeEmail) }
    if (currentUser.role !== 'admin') { taskData.assignee = currentUser.name; taskData.assigneeEmail = norm(currentUser.email) }
    if (editTask) { await updateDoc(doc(db, 'tasks', editTask.id), taskData); setEditTask(null) }
    else {
      await addDoc(collection(db, 'tasks'), { ...taskData, createdAt: serverTimestamp(), createdBy: norm(currentUser.email) })
      if (taskData.assigneeEmail && taskData.assigneeEmail !== norm(currentUser.email))
        await saveNotif(taskData.assigneeEmail, `משימה חדשה הוקצתה לך: "${taskData.title}"`, 'new')
    }
    setForm({ project: '', title: '', assignee: '', assigneeEmail: '', priority: 'medium', status: 'todo', dueDate: '', notes: '' })
    setShowForm(false)
  }

  const openEdit = task => {
    setEditTask(task)
    setForm({ project: task.project || '', title: task.title || '', assignee: task.assignee || '', assigneeEmail: task.assigneeEmail || '', priority: task.priority || 'medium', status: task.status || 'todo', dueDate: task.dueDate || '', notes: task.notes || '' })
    setShowForm(true)
  }

  const deleteTask = async id => {
    if (confirm('למחוק?')) {
      const task = tasks.find(t => t.id === id)
      if (task) await addDoc(collection(db, 'deleted_tasks'), { ...task, deletedAt: Date.now(), deletedBy: norm(currentUser.email) })
      await deleteDoc(doc(db, 'tasks', id))
    }
  }

  const deleteAllDone = async () => {
    const doneTasks = tasks.filter(t => t.status === 'done')
    if (!doneTasks.length) return alert('אין משימות שהושלמו')
    if (!confirm(`למחוק ${doneTasks.length} משימות שהושלמו? ניתן לשחזר עד 24 שעות`)) return
    for (const task of doneTasks) {
      await addDoc(collection(db, 'deleted_tasks'), { ...task, deletedAt: Date.now(), deletedBy: norm(currentUser.email) })
      await deleteDoc(doc(db, 'tasks', task.id))
    }
  }

  const restoreTask = async task => {
    const { id, deletedAt, deletedBy, ...taskData } = task
    await addDoc(collection(db, 'tasks'), taskData)
    await deleteDoc(doc(db, 'deleted_tasks', id))
    setDeletedTasks(prev => prev.filter(t => t.id !== id))
  }

  const changePassword = async () => {
    if (!newPassword.trim()) return
    await updateDoc(doc(db, 'users', selectedUserForPassword.id), { password: newPassword })
    setShowPasswordModal(false); setNewPassword('')
    alert(`סיסמה עודכנה עבור ${selectedUserForPassword.name}`)
  }

  const startResize = (e, col) => {
    e.preventDefault()
    const startX = e.clientX
    const startW = colWidths[col]
    const cols = Object.keys(colWidths)
    const nextCol = cols[cols.indexOf(col) + 1]
    const startNextW = colWidths[nextCol]
    setResizing(col)
    const onMove = e => setColWidths(prev => ({ ...prev, [col]: Math.max(60, startW + (startX - e.clientX)), ...(nextCol ? { [nextCol]: Math.max(60, startNextW - (startX - e.clientX)) } : {}) }))
    const onUp = () => { setResizing(null); window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp) }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }

  const autoFit = () => {
    if (!tableRef.current) return
    const totalW = tableRef.current.parentElement.offsetWidth
    const cols = Object.keys(colWidths)
    const mins = { project: 100, title: 150, assignee: 120, priority: 100, status: 110, dueDate: 110, notes: 150, actions: 80 }
    const minTotal = cols.reduce((s, c) => s + mins[c], 0)
    const ratio = totalW / minTotal
    const nw = {}
    cols.forEach(c => { nw[c] = Math.floor(mins[c] * ratio) })
    setColWidths(nw)
  }

    const exportToExcel = async () => {
    const ExcelJS = (await import('exceljs')).default
    const { saveAs } = await import('file-saver')

    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('משימות')

    ws.columns = [
      { header: 'פרויקט', key: 'project', width: 15 },
      { header: 'משימה', key: 'title', width: 30 },
      { header: 'אחראי', key: 'assignee', width: 15 },
      { header: 'עדיפות', key: 'priority', width: 12 },
      { header: 'סטטוס', key: 'status', width: 12 },
      { header: 'תאריך יעד', key: 'dueDate', width: 15 },
      { header: 'הערות', key: 'notes', width: 30 },
    ]

    displayedTasks.forEach(t => {
      ws.addRow({
        project: t.project || '',
        title: t.title || '',
        assignee: t.assignee || '',
        priority: PRIORITIES[t.priority]?.label || '',
        status: STATUSES[t.status]?.label || '',
        dueDate: t.dueDate || '',
        notes: t.notes || ''
      })
    })

    const assigneeList = users.map(u => u.name)
    const lastRow = displayedTasks.length + 100

    for (let i = 2; i <= lastRow; i++) {
      ws.getCell(`C${i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: [`"${assigneeList.join(',')}"`],
        showErrorMessage: true, errorTitle: 'שגיאה', error: 'בחר מהרשימה'
      }
      ws.getCell(`D${i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"גבוהה,בינונית,נמוכה"'],
        showErrorMessage: true, errorTitle: 'שגיאה', error: 'בחר מהרשימה'
      }
      ws.getCell(`E${i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"לביצוע,בביצוע,הושלם"'],
        showErrorMessage: true, errorTitle: 'שגיאה', error: 'בחר מהרשימה'
      }
    }

    const buffer = await wb.xlsx.writeBuffer()
    saveAs(new Blob([buffer]), `tasks_${today}.xlsx`)
  }

  const importFromExcel = async (e) => {
    const file = e.target.files[0]; if (!file) return
    const reader = new FileReader()
    reader.onload = async (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'binary' })
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
      const pMap = { 'גבוהה': 'high', 'בינונית': 'medium', 'נמוכה': 'low' }
      const sMap = { 'לביצוע': 'todo', 'בביצוע': 'inprogress', 'הושלם': 'done' }
      let added = 0, updated = 0, skipped = 0
      for (const row of rows) {
        const title = row['משימה']?.toString().trim(); if (!title) { skipped++; continue }
        const assigneeName = currentUser.role !== 'admin' ? currentUser.name : (row['אחראי']?.toString().trim() || '')
        const userMatch = currentUser.role !== 'admin' ? currentUser : users.find(u => u.name === assigneeName || norm(u.email) === norm(assigneeName))
        const assigneeEmail = currentUser.role !== 'admin' ? norm(currentUser.email) : norm(userMatch?.email || '')
        const project = row['פרויקט']?.toString().trim() || ''
        const priority = pMap[row['עדיפות']?.toString().trim()] || 'medium'
        const status = sMap[row['סטטוס']?.toString().trim()] || 'todo'
        const parseDate = val => {
  if (!val) return ''
    const s = String(val).trim()
    // מספר סידורי של Excel
    if (/^\d{4,5}$/.test(s)) {
      const d = new Date((parseInt(s) - 25569) * 86400 * 1000)
      const dd = String(d.getUTCDate()).padStart(2, '0')
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
      const yyyy = d.getUTCFullYear()
      return `${yyyy}-${mm}-${dd}`
    }
    // פורמט DD/MM/YYYY או DD-MM-YYYY
    const parts = s.split(/[\/\-\.]/)
    if (parts.length === 3) {
      const [a, b, c] = parts
      if (c.length === 4) return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`
      if (a.length === 4) return `${a}-${b.padStart(2,'0')}-${c.padStart(2,'0')}`
    }
    return s
  }
  const dueDate = parseDate(row['תאריך יעד'])
        const notes = row['הערות']?.toString().trim() || ''
        const existing = tasks.find(t => t.title?.trim() === title && norm(t.assigneeEmail) === assigneeEmail && (t.project?.trim() || '') === project)
        if (existing) {
          const changes = {}
          if (existing.priority !== priority) changes.priority = priority
          if (existing.status !== status) changes.status = status
          if ((existing.dueDate || '') !== dueDate) changes.dueDate = dueDate
          if ((existing.notes || '') !== notes) changes.notes = notes
          if (Object.keys(changes).length > 0) { await updateDoc(doc(db, 'tasks', existing.id), changes); updated++ } else skipped++
        } else {
          await addDoc(collection(db, 'tasks'), { project, title, assignee: assigneeName, assigneeEmail, priority, status, dueDate, notes, createdAt: serverTimestamp(), createdBy: norm(currentUser.email) })
          if (assigneeEmail && assigneeEmail !== norm(currentUser.email)) await saveNotif(assigneeEmail, `משימה חדשה הוקצתה לך: "${title}"`, 'new')
          added++
        }
      }
      alert(`✅ ייבוא הושלם!\nחדשות: ${added} | עודכנו: ${updated} | ללא שינוי: ${skipped}`)
    }
    reader.readAsBinaryString(file); e.target.value = ''
  }

  const isOverdue = d => d && new Date(d) < new Date()
  const isTomorrow = d => { if (!d) return false; const tom = new Date(); tom.setDate(tom.getDate() + 1); return d === tom.toISOString().split('T')[0] }
  const isAdmin = currentUser?.role === 'admin'
  const unreadCount = notifications.filter(n => !n.read).length
  const filteredByUser = selectedUser ? tasks.filter(t => norm(t.assigneeEmail) === norm(selectedUser.email)) : tasks

  const availableStatuses = [...new Set(filteredByUser.filter(t => (filterPriority === 'all' || t.priority === filterPriority) && (filterProject === 'all' || t.project === filterProject)).map(t => t.status))]
  const availablePriorities = [...new Set(filteredByUser.filter(t => (filterStatus === 'all' || t.status === filterStatus) && (filterProject === 'all' || t.project === filterProject)).map(t => t.priority))]
  const availableProjects = [...new Set(filteredByUser.filter(t => (filterStatus === 'all' || t.status === filterStatus) && (filterPriority === 'all' || t.priority === filterPriority)).map(t => t.project).filter(Boolean))].sort()

  const displayedTasks = filteredByUser
    .filter(t => filterStatus === 'all' || t.status === filterStatus)
    .filter(t => filterPriority === 'all' || t.priority === filterPriority)
    .filter(t => filterProject === 'all' || t.project === filterProject)
    .filter(t => !search || [t.title, t.project, t.assignee, t.notes].some(v => v?.includes(search)))
    .sort((a, b) => {
      if (!sortField) return 0
      let va = a[sortField] || '', vb = b[sortField] || ''
      if (sortField === 'priority') { va = PRIORITIES[a.priority]?.order || 0; vb = PRIORITIES[b.priority]?.order || 0 }
      if (sortField === 'dueDate') { va = va || '9999'; vb = vb || '9999' }
      return sortDir === 'asc' ? (va > vb ? 1 : -1) : (va < vb ? 1 : -1)
    })

  const bt = selectedUser ? filteredByUser : tasks
  const bi = { total: bt.length, todo: bt.filter(t => t.status === 'todo').length, inprogress: bt.filter(t => t.status === 'inprogress').length, done: bt.filter(t => t.status === 'done').length, overdue: bt.filter(t => isOverdue(t.dueDate) && t.status !== 'done').length, rate: bt.length ? Math.round(bt.filter(t => t.status === 'done').length / bt.length * 100) : 0 }

  const inputStyle = { width: '100%', padding: '10px 12px', border: '1.5px solid #e2e8f0', borderRadius: 10, boxSizing: 'border-box', fontSize: 14, outline: 'none', fontFamily: 'inherit', color: '#0f172a', background: 'white' }
  const btnPrimary = { padding: '10px 20px', background: 'linear-gradient(135deg, #10b981, #0ea5e9)', color: 'white', border: 'none', borderRadius: 10, cursor: 'pointer', fontWeight: 700, fontSize: 14, fontFamily: 'inherit', boxShadow: '0 2px 8px rgba(16,185,129,0.3)' }
  const btnSecondary = { padding: '10px 20px', background: 'white', border: '1.5px solid #e2e8f0', borderRadius: 10, cursor: 'pointer', fontSize: 14, fontFamily: 'inherit', color: '#475569' }
  const selectStyle = { padding: '8px 12px', border: '1.5px solid #e2e8f0', borderRadius: 10, fontSize: 13, cursor: 'pointer', background: 'white', color: '#0f172a', outline: 'none', fontFamily: 'inherit' }

  const SortBtn = ({ field, label }) => (
    <span style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' }}
      onClick={() => { if (sortField === field) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortField(field); setSortDir('asc') } }}>
      {label}
      <span style={{ display: 'flex', flexDirection: 'column', gap: 1, marginTop: 1 }}>
        <span style={{ fontSize: 9, lineHeight: 1, color: sortField === field && sortDir === 'asc' ? '#10b981' : '#d1d5db', fontWeight: 700 }}>▲</span>
        <span style={{ fontSize: 9, lineHeight: 1, color: sortField === field && sortDir === 'desc' ? '#10b981' : '#d1d5db', fontWeight: 700 }}>▼</span>
      </span>
    </span>
  )

  if (!currentUser) return (
    <div style={{ minHeight: '100vh', width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #10b981 0%, #0ea5e9 100%)' }}>
      <div style={{ background: 'white', padding: '2.5rem', borderRadius: 24, boxShadow: '0 25px 60px rgba(0,0,0,0.15)', width: 400 }}>
        <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
          <div style={{ width: 64, height: 64, borderRadius: 18, background: 'linear-gradient(135deg, #10b981, #0ea5e9)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 28, margin: '0 auto 1rem', boxShadow: '0 8px 20px rgba(16,185,129,0.3)' }}>🗂️</div>
          <h2 style={{ margin: '0 0 0.25rem', color: '#0f172a', fontSize: 22, fontWeight: 700 }}>Task Board</h2>
          <p style={{ margin: 0, color: '#94a3b8', fontSize: 13 }}>נס טכנולוגיות</p>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <input placeholder="אימייל" value={email} onChange={e => setEmail(e.target.value)} style={{ ...inputStyle, direction: 'ltr' }} />
          <input placeholder="סיסמה" type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === 'Enter' && login()} style={{ ...inputStyle, direction: 'ltr' }} />
          {error && <p style={{ color: '#dc2626', margin: 0, fontSize: 13, textAlign: 'center', background: '#fef2f2', padding: '8px 12px', borderRadius: 8 }}>{error}</p>}
          <button onClick={login} style={{ ...btnPrimary, padding: '12px', fontSize: 15, marginTop: 4 }}>כניסה למערכת</button>
        </div>
      </div>
    </div>
  )

  return (
    <div style={{ minHeight: '100vh', width: '100%', background: '#f8fafc', fontFamily: "'Segoe UI', system-ui, sans-serif", direction: 'rtl' }}>

      <div style={{ background: 'white', padding: '0.85rem 2rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #e2e8f0', boxShadow: '0 1px 8px rgba(0,0,0,0.06)', position: 'sticky', top: 0, zIndex: 50 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <div style={{ width: 38, height: 38, borderRadius: 10, background: 'linear-gradient(135deg, #10b981, #0ea5e9)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, boxShadow: '0 2px 8px rgba(16,185,129,0.3)' }}>🗂️</div>
          <div>
            <div style={{ color: '#0f172a', fontWeight: 700, fontSize: 16, lineHeight: 1 }}>Task Board</div>
            <div style={{ color: '#94a3b8', fontSize: 11, marginTop: 2 }}>נס טכנולוגיות {isAdmin && '· 👑 מנהל'}</div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          {isAdmin && <button onClick={() => setShowDashboard(true)} style={{ ...btnSecondary, padding: '7px 14px', fontSize: 13 }}>📊 Dashboard</button>}
          <button onClick={exportToExcel} style={{ ...btnSecondary, padding: '7px 14px', fontSize: 13 }}>📥 Excel</button>
          <button onClick={() => fileInputRef.current.click()} style={{ ...btnSecondary, padding: '7px 14px', fontSize: 13 }}>📤 ייבוא</button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={importFromExcel} style={{ display: 'none' }} />
          <button onClick={() => { setEditTask(null); setForm({ project: '', title: '', assignee: '', assigneeEmail: '', priority: 'medium', status: 'todo', dueDate: '', notes: '' }); setShowForm(true) }} style={btnPrimary}>
            + משימה חדשה
          </button>

          <div style={{ position: 'relative' }}>
            <button onClick={toggleNotifications} style={{ width: 38, height: 38, borderRadius: 10, background: '#f8fafc', border: '1.5px solid #e2e8f0', cursor: 'pointer', fontSize: 16, display: 'flex', alignItems: 'center', justifyContent: 'center', position: 'relative' }}>
              🔔
              {unreadCount > 0 && <span style={{ position: 'absolute', top: -4, right: -4, background: '#ef4444', color: 'white', borderRadius: '50%', width: 17, height: 17, fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700 }}>{unreadCount}</span>}
            </button>
            {showNotifications && (
              <div style={{ position: 'absolute', top: 46, left: 0, background: 'white', borderRadius: 16, boxShadow: '0 10px 40px rgba(0,0,0,0.15)', width: 360, zIndex: 200, overflow: 'hidden', border: '1px solid #e2e8f0' }}>
                <div style={{ padding: '14px 16px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontWeight: 700, color: '#0f172a', fontSize: 14 }}>🔔 התראות</span>
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    {notifications.some(n => !n.read) && (
                      <button onClick={markAllRead} style={{ fontSize: 11, color: '#2563eb', background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 6, padding: '3px 8px', cursor: 'pointer', fontWeight: 600 }}>✓ סמן הכל כנקרא</button>
                    )}
                    {notifications.length > 0 && (
                      <button onClick={clearAllNotifications} style={{ fontSize: 11, color: '#dc2626', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: 6, padding: '3px 8px', cursor: 'pointer', fontWeight: 600 }}>🗑️ נקה הכל</button>
                    )}
                    <button onClick={() => setShowNotifications(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: 16 }}>✕</button>
                  </div>
                </div>
                <div style={{ maxHeight: 400, overflowY: 'auto' }}>
                  {notifications.length === 0
                    ? <div style={{ padding: '2rem', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>אין התראות 🎉</div>
                    : notifications.map(n => (
                      <div key={n.id} onClick={() => markOneRead(n)}
                        style={{ padding: '12px 16px', borderBottom: '1px solid #f8fafc', background: n.read ? 'white' : '#f0fdf4', display: 'flex', gap: 10, cursor: 'pointer', transition: 'background 0.1s' }}
                        onMouseEnter={e => e.currentTarget.style.background = n.read ? '#f8fafc' : '#dcfce7'}
                        onMouseLeave={e => e.currentTarget.style.background = n.read ? 'white' : '#f0fdf4'}>
                        <span style={{ fontSize: 15, marginTop: 1 }}>{n.type === 'new' ? '✨' : n.type === 'status' ? '🔄' : n.type === 'priority' ? '⚡' : '📅'}</span>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 13, color: n.read ? '#94a3b8' : '#0f172a', lineHeight: 1.4 }}>{n.message}</div>
                          <div style={{ fontSize: 11, color: '#cbd5e1', marginTop: 3 }}>{n.createdAt?.seconds ? new Date(n.createdAt.seconds * 1000).toLocaleString('he-IL') : ''}</div>
                        </div>
                        {!n.read && <span style={{ width: 7, height: 7, borderRadius: '50%', background: '#10b981', flexShrink: 0, marginTop: 4 }} />}
                      </div>
                    ))}
                </div>
              </div>
            )}
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#f8fafc', borderRadius: 10, padding: '6px 12px', border: '1.5px solid #e2e8f0' }}>
            <Avatar name={currentUser.name} size={26} />
            <span style={{ color: '#0f172a', fontSize: 13, fontWeight: 600 }}>{currentUser.name}</span>
          </div>
          <button onClick={logout} style={{ ...btnSecondary, padding: '7px 14px', fontSize: 13 }}>יציאה</button>
        </div>
      </div>

      {isAdmin && (
        <>
          <div style={{ background: 'white', padding: '1.1rem 2rem', display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 12, borderBottom: '1px solid #e2e8f0' }}>
            <StatCard label="סה״כ משימות" value={bi.total} color="#0ea5e9" bg="#eff6ff" icon="📋" />
            <StatCard label="לביצוע" value={bi.todo} color="#64748b" bg="#f8fafc" icon="⭕" />
            <StatCard label="בביצוע" value={bi.inprogress} color="#2563eb" bg="#eff6ff" icon="⚡" />
            <StatCard label="הושלמו" value={bi.done} color="#16a34a" bg="#f0fdf4" icon="✅" />
            <StatCard label="חורגות" value={bi.overdue} color="#dc2626" bg="#fef2f2" icon="⚠️" />
            <StatCard label="אחוז ביצוע" value={`${bi.rate}%`} color="#10b981" bg="#f0fdf4" icon="📊" />
          </div>
          <div style={{ background: 'white', borderBottom: '1px solid #e2e8f0', padding: '0.65rem 2rem', display: 'flex', gap: 8, alignItems: 'center', overflowX: 'auto' }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: '#94a3b8', whiteSpace: 'nowrap' }}>👥 עובד:</span>
            <button onClick={() => setSelectedUser(null)} style={{ padding: '5px 14px', borderRadius: 999, border: `1.5px solid ${!selectedUser ? '#10b981' : '#e2e8f0'}`, cursor: 'pointer', fontWeight: 600, fontSize: 13, background: !selectedUser ? '#f0fdf4' : 'white', color: !selectedUser ? '#10b981' : '#64748b' }}>כולם</button>
            {users.map(u => (
              <div key={u.id} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <button onClick={() => setSelectedUser(selectedUser?.id === u.id ? null : u)}
                  style={{ padding: '5px 12px', borderRadius: 999, border: `1.5px solid ${selectedUser?.id === u.id ? '#10b981' : '#e2e8f0'}`, cursor: 'pointer', fontWeight: 500, fontSize: 13, background: selectedUser?.id === u.id ? '#f0fdf4' : 'white', color: selectedUser?.id === u.id ? '#10b981' : '#475569', display: 'flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap' }}>
                  <Avatar name={u.name} size={18} />{u.name}{u.role === 'admin' && ' 👑'}
                  <span style={{ background: '#f1f5f9', borderRadius: 999, padding: '1px 6px', fontSize: 11, fontWeight: 700, color: '#64748b' }}>{tasks.filter(t => norm(t.assigneeEmail) === norm(u.email)).length}</span>
                </button>
                {u.role !== 'admin' && <button onClick={() => { setSelectedUserForPassword(u); setShowPasswordModal(true) }} style={{ width: 26, height: 26, borderRadius: 7, background: '#fffbeb', color: '#d97706', border: '1px solid #fde68a', cursor: 'pointer', fontSize: 12, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>🔑</button>}
              </div>
            ))}
          </div>
        </>
      )}

      <div style={{ padding: '1.25rem 2rem' }}>
        <div style={{ background: 'white', borderRadius: 14, border: '1px solid #e2e8f0', padding: '1rem 1.25rem', marginBottom: '1.25rem', boxShadow: '0 1px 4px rgba(0,0,0,0.04)' }}>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <input placeholder="🔍 חיפוש..." value={search} onChange={e => setSearch(e.target.value)} style={{ ...inputStyle, width: 220, padding: '8px 12px', flex: 'none' }} />
            <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={selectStyle}>
              <option value="all">📋 כל הסטטוסים</option>
              {Object.entries(STATUSES).filter(([k]) => availableStatuses.includes(k)).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
            <select value={filterPriority} onChange={e => setFilterPriority(e.target.value)} style={selectStyle}>
              <option value="all">⚡ כל העדיפויות</option>
              {Object.entries(PRIORITIES).filter(([k]) => availablePriorities.includes(k)).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
            <select value={filterProject} onChange={e => setFilterProject(e.target.value)} style={selectStyle}>
              <option value="all">📁 כל הפרויקטים</option>
              {availableProjects.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
            <button onClick={autoFit} style={{ ...btnSecondary, padding: '8px 12px', fontSize: 12 }}>↔️ AutoFit</button>
            {isAdmin && <>
              <button onClick={deleteAllDone} style={{ padding: '8px 12px', background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>🗑️ מחק הושלמו</button>
              {deletedTasks.length > 0 && <button onClick={() => setShowDeleted(true)} style={{ padding: '8px 12px', background: '#f0fdf4', color: '#16a34a', border: '1px solid #bbf7d0', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>♻️ שחזור ({deletedTasks.length})</button>}
            </>}
            {(search || filterStatus !== 'all' || filterPriority !== 'all' || filterProject !== 'all' || sortField) && (
              <button onClick={() => { setSearch(''); setFilterStatus('all'); setFilterPriority('all'); setFilterProject('all'); setSortField('') }}
                style={{ padding: '8px 12px', background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>✕ נקה</button>
            )}
            <span style={{ color: '#94a3b8', fontSize: 13, marginRight: 'auto' }}>{displayedTasks.length} משימות</span>
          </div>
        </div>

        <div style={{ background: 'white', borderRadius: 16, overflow: 'auto', border: '1px solid #e2e8f0', boxShadow: '0 1px 8px rgba(0,0,0,0.05)' }}>
          <table ref={tableRef} style={{ borderCollapse: 'collapse', fontSize: 14, tableLayout: 'fixed', width: '100%', minWidth: Object.values(colWidths).reduce((a, b) => a + b, 0) }}>
            <colgroup>{Object.entries(colWidths).map(([col, w]) => <col key={col} style={{ width: w }} />)}</colgroup>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0' }}>
                {[['project','פרויקט'],['title','משימה'],['assignee','אחראי'],['priority','עדיפות'],['status','סטטוס'],['dueDate','תאריך יעד'],['notes','הערות'],['actions','']].map(([field, label]) => (
                  <th key={field} style={{ padding: '13px 16px', textAlign: 'right', color: '#374151', fontWeight: 700, fontSize: 13, position: 'relative', borderLeft: '1px solid #f1f5f9' }}>
                    {field !== 'actions' && field !== 'notes' ? <SortBtn field={field} label={label} /> : <span style={{ whiteSpace: 'nowrap' }}>{label}</span>}
                    {field !== 'actions' && <span onMouseDown={e => startResize(e, field)} style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: 5, cursor: 'col-resize', background: resizing === field ? '#10b981' : 'transparent', zIndex: 1 }} />}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {displayedTasks.length === 0 && (
                <tr><td colSpan={8} style={{ textAlign: 'center', padding: '3.5rem', color: '#cbd5e1' }}>
                  <div style={{ fontSize: 32, marginBottom: 8 }}>📭</div>
                  <div style={{ fontSize: 14 }}>אין משימות להצגה</div>
                </td></tr>
              )}
              {displayedTasks.map((task, i) => {
                const overdue = isOverdue(task.dueDate) && task.status !== 'done'
                const tomorrow = isTomorrow(task.dueDate) && task.status !== 'done'
                return (
                  <tr key={task.id}
                    style={{ borderBottom: '1px solid #f1f5f9', background: overdue ? '#fff5f5' : i % 2 === 0 ? 'white' : '#fafbfc', transition: 'background 0.1s', borderRight: overdue ? '3px solid #ef4444' : tomorrow ? '3px solid #f59e0b' : '3px solid transparent' }}
                    onMouseEnter={e => e.currentTarget.style.background = '#f0fdf4'}
                    onMouseLeave={e => e.currentTarget.style.background = overdue ? '#fff5f5' : i % 2 === 0 ? 'white' : '#fafbfc'}>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      {task.project ? <span style={{ background: '#eff6ff', color: '#1d4ed8', padding: '3px 10px', borderRadius: 6, fontSize: 12, fontWeight: 700 }}>{task.project}</span> : <span style={{ color: '#e2e8f0' }}>—</span>}
                    </td>
                    <td style={{ padding: '13px 16px', fontWeight: 600, color: '#0f172a', fontSize: 14, borderLeft: '1px solid #f1f5f9', verticalAlign: 'top', whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{task.title}</td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      {task.assignee ? <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}><Avatar name={task.assignee} size={26} /><span style={{ color: '#0f172a', fontSize: 13, fontWeight: 500 }}>{task.assignee}</span></div> : <span style={{ color: '#e2e8f0' }}>—</span>}
                    </td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      <Badge {...(PRIORITIES[task.priority] || { label: task.priority, color: '#f8fafc', text: '#475569', dot: '#94a3b8' })} />
                    </td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      <select value={task.status} onChange={e => updateDoc(doc(db, 'tasks', task.id), { status: e.target.value })}
                        style={{ border: `1.5px solid ${STATUSES[task.status]?.dot || '#94a3b8'}50`, background: STATUSES[task.status]?.color || '#f8fafc', color: STATUSES[task.status]?.text || '#475569', padding: '5px 10px', borderRadius: 999, fontSize: 12, fontWeight: 700, cursor: 'pointer', outline: 'none', fontFamily: 'inherit' }}>
                        {Object.entries(STATUSES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      {task.dueDate
                        ? <span style={{ color: overdue ? '#dc2626' : tomorrow ? '#d97706' : '#0f172a', fontWeight: overdue || tomorrow ? 700 : 500, fontSize: 13, display: 'flex', alignItems: 'center', gap: 4 }}>
                            {overdue && '⚠️'}{tomorrow && '⏰'} {task.dueDate}
                          </span>
                        : <span style={{ color: '#e2e8f0' }}>—</span>}
                    </td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}><Tooltip text={task.notes} /></td>
                    <td style={{ padding: '13px 16px', borderLeft: '1px solid #f1f5f9', verticalAlign: 'top' }}>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button onClick={() => openEdit(task)} style={{ padding: '5px 10px', background: '#eff6ff', color: '#2563eb', border: '1px solid #bfdbfe', borderRadius: 7, cursor: 'pointer', fontSize: 13, fontWeight: 600 }}>✏️</button>
                        {isAdmin && <button onClick={() => deleteTask(task.id)} style={{ padding: '5px 10px', background: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', borderRadius: 7, cursor: 'pointer', fontSize: 13, fontWeight: 600 }}>🗑️</button>}
                      </div>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </div>

      {showDashboard && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, backdropFilter: 'blur(4px)' }}>
          <div style={{ background: 'white', padding: '2rem', borderRadius: 20, width: 700, maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 25px 60px rgba(0,0,0,0.2)', border: '1px solid #e2e8f0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <h3 style={{ margin: 0, color: '#0f172a', fontSize: 18, fontWeight: 700 }}>📊 Dashboard — התקדמות לפי עובד</h3>
              <button onClick={() => setShowDashboard(false)} style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, width: 32, height: 32, cursor: 'pointer', fontSize: 16, color: '#64748b' }}>✕</button>
            </div>
            {users.filter(u => u.role !== 'admin').map(u => {
              const ut = tasks.filter(t => norm(t.assigneeEmail) === norm(u.email))
              const done = ut.filter(t => t.status === 'done').length
              const inprog = ut.filter(t => t.status === 'inprogress').length
              const todo = ut.filter(t => t.status === 'todo').length
              const over = ut.filter(t => isOverdue(t.dueDate) && t.status !== 'done').length
              const rate = ut.length ? Math.round(done / ut.length * 100) : 0
              return (
                <div key={u.id} style={{ marginBottom: '1.5rem', background: '#f8fafc', borderRadius: 12, padding: '1rem 1.25rem', border: '1px solid #e2e8f0' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
                    <Avatar name={u.name} size={34} />
                    <div>
                      <div style={{ fontWeight: 700, color: '#0f172a', fontSize: 15 }}>{u.name}</div>
                      <div style={{ fontSize: 12, color: '#94a3b8' }}>{ut.length} משימות סה״כ</div>
                    </div>
                    <div style={{ marginRight: 'auto', fontWeight: 700, fontSize: 18, color: rate >= 75 ? '#16a34a' : rate >= 40 ? '#d97706' : '#dc2626' }}>{rate}%</div>
                  </div>
                  <div style={{ display: 'flex', gap: 6, marginBottom: 10 }}>
                    {[{ label: 'לביצוע', val: todo, color: '#94a3b8' }, { label: 'בביצוע', val: inprog, color: '#3b82f6' }, { label: 'הושלם', val: done, color: '#22c55e' }, { label: 'חורגות', val: over, color: '#ef4444' }].map(s => (
                      <div key={s.label} style={{ flex: 1, textAlign: 'center', background: 'white', borderRadius: 8, padding: '6px', border: '1px solid #e2e8f0' }}>
                        <div style={{ fontSize: 18, fontWeight: 700, color: s.color }}>{s.val}</div>
                        <div style={{ fontSize: 11, color: '#94a3b8' }}>{s.label}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ background: '#e2e8f0', borderRadius: 999, height: 10, overflow: 'hidden' }}>
                    <div style={{ width: `${rate}%`, height: '100%', background: rate >= 75 ? 'linear-gradient(90deg, #10b981, #22c55e)' : rate >= 40 ? 'linear-gradient(90deg, #f59e0b, #fbbf24)' : 'linear-gradient(90deg, #ef4444, #f87171)', borderRadius: 999, transition: 'width 0.5s' }} />
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      )}

      {showDeleted && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, backdropFilter: 'blur(4px)' }}>
          <div style={{ background: 'white', padding: '2rem', borderRadius: 20, width: 560, maxHeight: '80vh', overflowY: 'auto', boxShadow: '0 25px 60px rgba(0,0,0,0.2)', border: '1px solid #e2e8f0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <h3 style={{ margin: 0, color: '#0f172a', fontSize: 18, fontWeight: 700 }}>♻️ שחזור משימות</h3>
              <button onClick={() => setShowDeleted(false)} style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, width: 32, height: 32, cursor: 'pointer', fontSize: 16, color: '#64748b' }}>✕</button>
            </div>
            <p style={{ color: '#94a3b8', fontSize: 13, marginBottom: '1rem' }}>משימות שנמחקו ב-24 השעות האחרונות</p>
            {deletedTasks.length === 0
              ? <div style={{ textAlign: 'center', color: '#94a3b8', padding: '2rem' }}>אין משימות לשחזור</div>
              : deletedTasks.map(t => (
                <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px', borderRadius: 10, border: '1px solid #e2e8f0', marginBottom: 8, background: '#f8fafc' }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, color: '#0f172a', fontSize: 14 }}>{t.title}</div>
                    <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 2 }}>
                      {t.assignee && `👤 ${t.assignee} · `}{t.project && `📁 ${t.project} · `}נמחק לפני {Math.round((Date.now() - t.deletedAt) / 60000)} דקות
                    </div>
                  </div>
                  <button onClick={() => restoreTask(t)} style={{ padding: '6px 14px', background: '#f0fdf4', color: '#16a34a', border: '1px solid #bbf7d0', borderRadius: 8, cursor: 'pointer', fontSize: 13, fontWeight: 600 }}>שחזר</button>
                </div>
              ))}
          </div>
        </div>
      )}

      {showForm && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, backdropFilter: 'blur(4px)' }}>
          <div style={{ background: 'white', padding: '2rem', borderRadius: 20, width: 520, boxShadow: '0 25px 60px rgba(0,0,0,0.2)', maxHeight: '90vh', overflowY: 'auto', border: '1px solid #e2e8f0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <h3 style={{ margin: 0, color: '#0f172a', fontSize: 18, fontWeight: 700 }}>{editTask ? '✏️ עריכת משימה' : '✨ משימה חדשה'}</h3>
              <button onClick={() => { setShowForm(false); setEditTask(null) }} style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, width: 32, height: 32, cursor: 'pointer', fontSize: 16, color: '#64748b' }}>✕</button>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              {[['project', 'פרויקט', 'text'], ['title', 'משימה *', 'text'], ['dueDate', 'תאריך יעד', 'date']].map(([key, label, type]) => (
                <div key={key}>
                  <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>{label}</label>
                  <input type={type} value={form[key]} min={key === 'dueDate' ? today : undefined} onChange={e => setForm({ ...form, [key]: e.target.value })} style={inputStyle} />
                </div>
              ))}
              {isAdmin ? (
                <div>
                  <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>שייך לעובד</label>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 200, overflowY: 'auto', border: '1.5px solid #e2e8f0', borderRadius: 10, padding: 6 }}>
                    {users.map(u => (
                      <div key={u.id} onClick={() => setForm({ ...form, assigneeEmail: norm(u.email), assignee: u.name })}
                        style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 10px', borderRadius: 8, cursor: 'pointer', background: norm(form.assigneeEmail) === norm(u.email) ? '#f0fdf4' : 'white', border: `1.5px solid ${norm(form.assigneeEmail) === norm(u.email) ? '#10b981' : 'transparent'}`, transition: 'all 0.1s' }}
                        onMouseEnter={e => { if (norm(form.assigneeEmail) !== norm(u.email)) e.currentTarget.style.background = '#f8fafc' }}
                        onMouseLeave={e => { if (norm(form.assigneeEmail) !== norm(u.email)) e.currentTarget.style.background = 'white' }}>
                        <Avatar name={u.name} size={30} />
                        <div>
                          <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{u.name}{u.role === 'admin' ? ' 👑' : ''}</div>
                          <div style={{ fontSize: 11, color: '#94a3b8' }}>{u.email}</div>
                        </div>
                        {norm(form.assigneeEmail) === norm(u.email) && <span style={{ marginRight: 'auto', color: '#10b981', fontWeight: 700, fontSize: 16 }}>✓</span>}
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
                <div>
                  <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>אחראי</label>
                  <input value={currentUser.name} disabled style={{ ...inputStyle, background: '#f8fafc', color: '#94a3b8' }} />
                </div>
              )}
              <div>
                <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>עדיפות</label>
                <div style={{ display: 'flex', gap: 8 }}>
                  {Object.entries(PRIORITIES).map(([k, v]) => (
                    <div key={k} onClick={() => setForm({ ...form, priority: k })}
                      style={{ flex: 1, padding: '8px', borderRadius: 8, border: `1.5px solid ${form.priority === k ? v.dot : '#e2e8f0'}`, cursor: 'pointer', textAlign: 'center', background: form.priority === k ? v.color : 'white', transition: 'all 0.1s' }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: form.priority === k ? v.text : '#94a3b8' }}>{v.label}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div>
                <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>סטטוס</label>
                <div style={{ display: 'flex', gap: 6 }}>
                  {Object.entries(STATUSES).map(([k, v]) => (
                    <div key={k} onClick={() => setForm({ ...form, status: k })}
                      style={{ flex: 1, padding: '8px 4px', borderRadius: 8, border: `1.5px solid ${form.status === k ? v.dot : '#e2e8f0'}`, cursor: 'pointer', textAlign: 'center', background: form.status === k ? v.color : 'white', transition: 'all 0.1s' }}>
                      <div style={{ fontSize: 11, fontWeight: 700, color: form.status === k ? v.text : '#94a3b8' }}>{v.label}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            <div style={{ marginTop: 12 }}>
              <label style={{ display: 'block', marginBottom: 5, fontSize: 12, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.06em' }}>הערות</label>
              <textarea value={form.notes} onChange={e => setForm({ ...form, notes: e.target.value })} rows={3} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
            </div>
            <div style={{ display: 'flex', gap: 10, marginTop: '1.5rem' }}>
              <button onClick={saveTask} style={{ ...btnPrimary, flex: 1, padding: '12px' }}>{editTask ? 'שמור שינויים' : 'הוסף משימה'}</button>
              <button onClick={() => { setShowForm(false); setEditTask(null) }} style={{ ...btnSecondary, flex: 1, padding: '12px' }}>ביטול</button>
            </div>
          </div>
        </div>
      )}

      {showPasswordModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100, backdropFilter: 'blur(4px)' }}>
          <div style={{ background: 'white', padding: '2rem', borderRadius: 20, width: 380, boxShadow: '0 25px 60px rgba(0,0,0,0.2)', border: '1px solid #e2e8f0' }}>
            <h3 style={{ margin: '0 0 0.5rem', color: '#0f172a' }}>🔑 שינוי סיסמה</h3>
            <p style={{ color: '#64748b', fontSize: 14, marginBottom: '1.25rem' }}>עובד: <strong style={{ color: '#0f172a' }}>{selectedUserForPassword?.name}</strong></p>
            <input placeholder="סיסמה חדשה" type="text" value={newPassword} onChange={e => setNewPassword(e.target.value)} style={{ ...inputStyle, marginBottom: 16 }} />
            <div style={{ display: 'flex', gap: 10 }}>
              <button onClick={changePassword} style={{ ...btnPrimary, flex: 1, padding: '11px' }}>שמור</button>
              <button onClick={() => { setShowPasswordModal(false); setNewPassword('') }} style={{ ...btnSecondary, flex: 1, padding: '11px' }}>ביטול</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}