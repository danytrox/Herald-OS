import { shell } from 'electron'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import type { BatteryStatus, CalendarResult, ControlAction, DiskUsage, InstalledApp, NetworkStatus, ProcessInfo, RecentFile, StatusPanelId, StatusPanelState, SystemInfo, SystemStats } from '../../shared/ipc.ts'
import { findCodeEditor } from './editors.ts'
import { run } from './exec.ts'
import { normaliseFileQuery, rankFiles } from './find.ts'
import { type EditorTarget, type HostPlatform, HostNotSupported } from './types.ts'

/*
 * Windows host facts and actions for the shell. Reads come from Node's `os` where it is enough
 * (memory, CPU, network) and from PowerShell / WMI where it is not (processes, disks, battery, the
 * OS name); the installed-app list scans the Start Menu. Still placeholders on Windows: the quick
 * panels (`controlStatus` / `controlAction`), the calendar, application icons and per-process
 * owner names (Win32_Process.GetOwner is far too slow per process to ship here).
 */

// --- pure parsers (tested) -------------------------------------------------

interface WinProcessRow {
  ProcessId?: number
  ParentProcessId?: number
  Name?: string
  CommandLine?: string | null
  WorkingSetSize?: number | string
  UserModeTime?: number | string
  KernelModeTime?: number | string
  CreationDate?: string
}

interface WinDiskRow { DeviceID?: string; Size?: number | string; FreeSpace?: number | string }
interface WinBatteryRow { EstimatedChargeRemaining?: number | string; BatteryStatus?: number | string }
interface WinOsRow { Caption?: string; Version?: string }

/** `ConvertTo-Json` in Windows PowerShell 5.1 emits a bare object for one row and an array for many. */
function asArray<T>(json: unknown): T[] {
  if (Array.isArray(json)) {
    return json as T[]
  }

  return json === null || json === undefined ? [] : [json as T]
}

function clamp(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(high, value))
}

/** Parse `Get-CimInstance Win32_Process | Select ... | ConvertTo-Json`. */
export function parseWinProcesses(text: string, totalMemory = os.totalmem()): ProcessInfo[] {
  let rows: WinProcessRow[] = []

  try {
    rows = asArray<WinProcessRow>(JSON.parse(text))
  } catch {
    return []
  }

  const now = Date.now()

  return rows.flatMap(row => {
    if (!row || !Number.isFinite(Number(row.ProcessId))) {
      return []
    }

    // KernelModeTime + UserModeTime are 100 ns ticks; together they are the process's CPU time.
    const cpuSeconds = (Number(row.UserModeTime ?? 0) + Number(row.KernelModeTime ?? 0)) / 10_000_000
    // ConvertTo-Json renders the DateTime as ASP.NET's `\/Date(ms)\/`, not ISO 8601.
    const started = /\/Date\((\d+)\)\//.exec(row.CreationDate ?? '')?.[1] ?? Date.parse(row.CreationDate ?? '')
    const elapsedSeconds = Number.isFinite(Number(started)) && Number(started) > 0 ? (now - Number(started)) / 1000 : 0
    const rssBytes = Math.max(0, Number(row.WorkingSetSize ?? 0))
    const name = row.Name ?? ''

    return [{
      pid: Number(row.ProcessId),
      ppid: Number(row.ParentProcessId ?? 0),
      // Owner is empty: Win32_Process.GetOwner is a slow per-process call, not worth it per refresh.
      user: '',
      cpuPercent: elapsedSeconds > 0 ? Math.max(0, cpuSeconds / elapsedSeconds * 100) : 0,
      memPercent: totalMemory > 0 ? clamp(rssBytes / totalMemory * 100, 0, 100) : 0,
      rssBytes,
      command: row.CommandLine ?? name,
      name
    }]
  })
}

/** Parse `Get-CimInstance Win32_LogicalDisk | Select DeviceID,Size,FreeSpace | ConvertTo-Json`. */
export function parseWinDisks(text: string): DiskUsage[] {
  let rows: WinDiskRow[] = []

  try {
    rows = asArray<WinDiskRow>(JSON.parse(text))
  } catch {
    return []
  }

  return rows.flatMap(row => {
    const total = Number(row?.Size ?? 0)
    const free = Number(row?.FreeSpace ?? 0)

    if (!row?.DeviceID || !Number.isFinite(total) || total <= 0) {
      return []
    }

    return [{ mount: `${row.DeviceID}\\`, total, free, used: Math.max(0, total - free) }]
  })
}

/**
 * Parse `Get-CimInstance Win32_Battery | Select EstimatedChargeRemaining,BatteryStatus`.
 * BatteryStatus 1 = discharging; 2, 3 and 6-11 mean plugged in or charging.
 */
export function parseWinBattery(text: string): BatteryStatus {
  let rows: WinBatteryRow[] = []

  try {
    rows = asArray<WinBatteryRow>(JSON.parse(text))
  } catch {
    return { present: false }
  }

  const battery = rows.find(row => row && Number.isFinite(Number(row.EstimatedChargeRemaining)))

  if (!battery) {
    return { present: false }
  }

  const status = Number(battery.BatteryStatus)

  return {
    present: true,
    percent: clamp(Number(battery.EstimatedChargeRemaining), 0, 100),
    charging: [2, 3, 6, 7, 8, 9, 11].includes(status)
  }
}

/** Parse `Get-NetConnectionProfile | Select Name,InterfaceAlias`: the Wi-Fi profile's SSID, or null. */
export function parseNetWifi(text: string): NonNullable<NetworkStatus['wifi']> | null {
  let rows: Array<{ Name?: string; InterfaceAlias?: string }> = []

  try {
    rows = asArray(JSON.parse(text))
  } catch {
    return null
  }

  // `netsh wlan` localises its labels, but Get-NetConnectionProfile uses stable .NET property names.
  const wifi = rows.find(row => row && /wi-?fi/i.test(row.InterfaceAlias ?? ''))

  if (!wifi) {
    return null
  }

  const ssid = wifi.Name?.trim()

  return { connected: Boolean(ssid), ssid: ssid || undefined }
}

/** Parse `Get-CimInstance Win32_OperatingSystem | Select Caption,Version`. */
export function parseWinOs(text: string): { osName: string; osVersion: string } {
  let rows: WinOsRow[] = []

  try {
    rows = asArray<WinOsRow>(JSON.parse(text))
  } catch {
    return { osName: 'Windows', osVersion: os.release() }
  }

  const row = rows[0]

  return { osName: row?.Caption?.trim() || 'Windows', osVersion: row?.Version?.trim() || os.release() }
}

// --- helpers ---------------------------------------------------------------

/** Run a PowerShell one-liner and return its UTF-8 stdout, or '' when it fails. */
async function ps(script: string, timeoutMs = 8000): Promise<string> {
  const result = await run(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', `[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false); ${script}`],
    timeoutMs
  )

  return result.code === 0 ? result.stdout.replace(/^\uFEFF/, '') : ''
}

/** Start Menu shortcuts, user's first so their overrides win; recurse a few levels for grouped apps. */
const START_MENU_DIRS = [
  path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
  path.join(process.env.ProgramData || 'C:\\ProgramData', 'Microsoft', 'Windows', 'Start Menu', 'Programs')
]

/** Folders a recursive walk must never descend into. */
const NOISE_DIRS = new Set(['node_modules', '.git', 'venv', '.venv', '__pycache__', 'dist', 'build', 'appdata', 'application data', 'library'])
const RAW_IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp'])
const RAW_IMAGE_LIMIT = 8 * 1024 * 1024

function cpuSample(): { idle: number; total: number } {
  let idle = 0
  let total = 0

  for (const cpu of os.cpus()) {
    idle += cpu.times.idle
    total += cpu.times.user + cpu.times.nice + cpu.times.sys + cpu.times.irq + cpu.times.idle
  }

  return { idle, total }
}

// --- the platform ----------------------------------------------------------

export class Win32Platform implements HostPlatform {
  private previousCpu = cpuSample()
  private appsCache: { at: number; apps: InstalledApp[] } | null = null

  async systemInfo(): Promise<SystemInfo> {
    const cpus = os.cpus()
    const win = parseWinOs(await ps('Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version | ConvertTo-Json -Compress'))

    return {
      hostname: os.hostname(),
      platform: 'win32',
      arch: process.arch,
      osName: win.osName,
      osVersion: win.osVersion,
      cpuModel: cpus[0]?.model ?? 'unknown',
      cpuCount: cpus.length,
      totalMemory: os.totalmem(),
      userName: os.userInfo().username,
      homeDir: os.homedir()
    }
  }

  async sampleStats(): Promise<SystemStats> {
    const current = cpuSample()
    const idleDelta = current.idle - this.previousCpu.idle
    const totalDelta = current.total - this.previousCpu.total
    this.previousCpu = current
    const cpuPercent = totalDelta > 0 ? clamp((1 - idleDelta / totalDelta) * 100, 0, 100) : 0

    const memoryTotal = os.totalmem()
    const memoryFree = os.freemem()
    const [disks, battery] = await Promise.all([
      ps('Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3 or DriveType=2" | Select-Object DeviceID,Size,FreeSpace | ConvertTo-Json -Compress'),
      ps('Get-CimInstance Win32_Battery | Select-Object EstimatedChargeRemaining,BatteryStatus | ConvertTo-Json -Compress')
    ])

    return {
      sampledAt: Date.now(),
      cpuPercent,
      loadAverage: os.loadavg() as [number, number, number],
      memoryTotal,
      memoryUsed: memoryTotal - memoryFree,
      memoryFree,
      uptimeSeconds: os.uptime(),
      disks: parseWinDisks(disks),
      battery: parseWinBattery(battery)
    }
  }

  async listProcesses(sort: 'cpu' | 'memory', limit: number): Promise<ProcessInfo[]> {
    const text = await ps('Get-CimInstance Win32_Process | Select-Object ProcessId,ParentProcessId,Name,CommandLine,WorkingSetSize,UserModeTime,KernelModeTime,CreationDate | ConvertTo-Json -Compress', 12_000)
    const rows = parseWinProcesses(text)
    rows.sort((a, b) => (sort === 'cpu' ? b.cpuPercent - a.cpuPercent : b.rssBytes - a.rssBytes))

    return rows.slice(0, limit)
  }

  async listInstalledApps(): Promise<InstalledApp[]> {
    if (this.appsCache && Date.now() - this.appsCache.at < 60_000) {
      return this.appsCache.apps
    }

    const apps = new Map<string, InstalledApp>()

    const visit = async (dir: string, depth: number): Promise<void> => {
      let dirents: import('node:fs').Dirent[] = []

      try {
        dirents = await fs.readdir(dir, { withFileTypes: true })
      } catch {
        return
      }

      for (const dirent of dirents) {
        const full = path.join(dir, dirent.name)

        if (dirent.isDirectory()) {
          if (depth < 4 && !dirent.name.startsWith('.')) {
            await visit(full, depth + 1)
          }

          continue
        }

        if (!dirent.isFile() || path.extname(dirent.name).toLowerCase() !== '.lnk') {
          continue
        }

        const name = dirent.name.slice(0, -'.lnk'.length)

        if (!apps.has(name.toLowerCase())) {
          apps.set(name.toLowerCase(), { name, path: full })
        }
      }
    }

    for (const dir of START_MENU_DIRS) {
      await visit(dir, 0)
    }

    const sorted = [...apps.values()].sort((a, b) => a.name.localeCompare(b.name))
    this.appsCache = { at: Date.now(), apps: sorted }

    return sorted
  }

  /** Extracting an icon from a .lnk or .exe needs the shell's icon cache; not worth it yet. */
  appIcon(): Promise<Buffer | null> {
    return Promise.resolve(null)
  }

  async launchApp(appPath: string): Promise<void> {
    const error = await shell.openPath(appPath)

    if (error) {
      throw new Error(error)
    }
  }

  revealPath(targetPath: string): Promise<void> {
    shell.showItemInFolder(targetPath)

    return Promise.resolve()
  }

  async openIn(target: EditorTarget, targetPath: string): Promise<void> {
    switch (target) {
      case 'editor':
        // A .lnk cannot take a file argument; opening the editor itself is the best a shortcut gives.
        await this.launchApp(findCodeEditor(await this.listInstalledApps()).path)

        return
      case 'finder': {
        let dir = targetPath

        try {
          if (!(await fs.stat(targetPath)).isDirectory()) {
            dir = path.dirname(targetPath)
          }
        } catch {
          dir = path.dirname(targetPath)
        }

        await shell.openPath(dir)

        return
      }
      case 'terminal': {
        const stat = await fs.stat(targetPath).catch(() => null)
        const dir = stat?.isDirectory() ? targetPath : path.dirname(targetPath)
        const wt = await run('wt.exe', ['-d', dir], 4000)

        if (wt.code !== 0) {
          // Windows Terminal is not installed: fall back to a plain console.
          const cmd = await run('cmd.exe', ['/c', 'start', '', 'cmd.exe', '/k', `cd /d "${dir}"`], 4000)

          if (cmd.code !== 0) {
            throw new HostNotSupported('external terminal')
          }
        }

        return
      }
    }
  }

  async networkStatus(): Promise<NetworkStatus> {
    let iface: string | undefined
    let ipv4: string | undefined

    for (const [name, addresses] of Object.entries(os.networkInterfaces())) {
      const address = addresses?.find(a => a.family === 'IPv4' && !a.internal)

      if (address) {
        iface = name
        ipv4 = address.address

        break
      }
    }

    const wifi = parseNetWifi(await ps('Get-NetConnectionProfile | Select-Object Name,InterfaceAlias | ConvertTo-Json -Compress'))

    return { online: Boolean(iface), defaultInterface: iface, ipv4, wifi: wifi ?? undefined }
  }

  /** Windows has no bridge to the built-in calendar without an Outlook/Graph sign-in. */
  async calendarToday(): Promise<CalendarResult> {
    return { status: 'unavailable', events: [] }
  }

  /** No cheap \"recently used\" source on Windows: files under the usual folders touched in the last 3 days. */
  async recentFiles(limit: number): Promise<RecentFile[]> {
    const home = os.homedir()
    const oneDrive = path.join(home, 'OneDrive')
    const since = Date.now() - 3 * 24 * 60 * 60 * 1000
    const roots = ['Desktop', 'Documents', 'Downloads', 'Pictures', 'Music', 'Videos']
      .flatMap(folder => [path.join(home, folder), path.join(oneDrive, folder)])
    const rows: RecentFile[] = []

    const visit = async (dir: string, depth: number): Promise<void> => {
      let dirents: import('node:fs').Dirent[] = []

      try {
        dirents = await fs.readdir(dir, { withFileTypes: true })
      } catch {
        return
      }

      for (const dirent of dirents) {
        if (dirent.name.startsWith('.')) {
          continue
        }

        const full = path.join(dir, dirent.name)

        if (dirent.isDirectory()) {
          if (depth < 2 && !NOISE_DIRS.has(dirent.name.toLowerCase())) {
            await visit(full, depth + 1)
          }

          continue
        }

        if (!dirent.isFile()) {
          continue
        }

        try {
          const stat = await fs.stat(full)

          if (stat.mtimeMs >= since) {
            rows.push({ path: full, name: dirent.name, extension: path.extname(dirent.name).toLowerCase(), size: stat.size, modifiedAt: stat.mtimeMs, lastUsedAt: stat.mtimeMs, kind: 'file' })
          }
        } catch {
          // Vanished or unreadable.
        }
      }
    }

    for (const root of roots) {
      await visit(root, 1)
    }

    rows.sort((a, b) => b.lastUsedAt - a.lastUsedAt)

    return rows.slice(0, limit)
  }

  /** No Spotlight on Windows: a bounded recursive walk over the home folder, noise pruned. */
  async findFiles(query: string, limit: number): Promise<RecentFile[]> {
    const name = normaliseFileQuery(query)

    if (!name) {
      return []
    }

    const home = os.homedir()
    const pattern = name.toLowerCase()
    const rows: RecentFile[] = []

    const visit = async (dir: string, depth: number): Promise<void> => {
      if (rows.length >= 300) {
        return
      }

      let dirents: import('node:fs').Dirent[] = []

      try {
        dirents = await fs.readdir(dir, { withFileTypes: true })
      } catch {
        return
      }

      for (const dirent of dirents) {
        if (rows.length >= 300) {
          return
        }

        if (dirent.name.startsWith('.')) {
          continue
        }

        const full = path.join(dir, dirent.name)

        if (dirent.isDirectory()) {
          if (depth < 6 && !NOISE_DIRS.has(dirent.name.toLowerCase())) {
            await visit(full, depth + 1)
          }

          continue
        }

        if (!dirent.name.toLowerCase().includes(pattern)) {
          continue
        }

        try {
          const stat = await fs.stat(full)
          rows.push({ path: full, name: dirent.name, extension: path.extname(dirent.name).toLowerCase(), size: stat.size, modifiedAt: stat.mtimeMs, lastUsedAt: stat.atimeMs, kind: 'file' })
        } catch {
          // Vanished or unreadable.
        }
      }
    }

    await visit(home, 0)

    return rankFiles(name, rows, limit)
  }

  /** No QuickLook on Windows: return small image files verbatim, nothing for other kinds. */
  async thumbnail(filePath: string): Promise<Buffer | null> {
    let stat: import('node:fs').Stats

    try {
      stat = await fs.stat(filePath)
    } catch {
      return null
    }

    if (!stat.isFile() || !RAW_IMAGE_EXTENSIONS.has(path.extname(filePath).toLowerCase()) || stat.size > RAW_IMAGE_LIMIT) {
      return null
    }

    return fs.readFile(filePath).catch(() => null)
  }

  controlStatus(_panel: StatusPanelId): Promise<StatusPanelState> {
    throw new HostNotSupported('quick panels')
  }

  controlAction(_action: ControlAction): Promise<void> {
    throw new HostNotSupported('quick panels')
  }
}
