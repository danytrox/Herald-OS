import { describe, expect, it } from 'vitest'
import { parseNetWifi, parseWinBattery, parseWinDisks, parseWinOs, parseWinProcesses } from './win32.ts'

describe('parseWinProcesses', () => {
  it('parses a Win32_Process JSON array into process rows', () => {
    const started = Date.now() - 10_000
    const text = JSON.stringify([
      { ProcessId: 1234, ParentProcessId: 4, Name: 'chrome.exe', CommandLine: 'C:\\Program Files\\Chrome\\chrome.exe --flag', WorkingSetSize: '123456789', UserModeTime: '15000000', KernelModeTime: '0', CreationDate: `/Date(${started})/` }
    ])
    const rows = parseWinProcesses(text, 1_000_000_000)

    expect(rows).toHaveLength(1)
    expect(rows[0]).toMatchObject({
      pid: 1234,
      ppid: 4,
      user: '',
      name: 'chrome.exe',
      command: 'C:\\Program Files\\Chrome\\chrome.exe --flag',
      rssBytes: 123456789
    })
    // 15_000_000 ticks / 1e7 = 1.5 CPU-seconds over 10 s = 15%.
    expect(rows[0].cpuPercent).toBeCloseTo(15, 0)
    expect(rows[0].memPercent).toBeCloseTo(12.3456789, 3)
  })

  it('tolerates a single object (PowerShell ConvertTo-Json for one row) and drops junk', () => {
    const one = JSON.stringify({ ProcessId: 7, Name: 'System Idle Process', CommandLine: null, WorkingSetSize: '8192', UserModeTime: '0', KernelModeTime: '0' })
    const rows = parseWinProcesses(one, 1_000_000_000)

    expect(rows).toHaveLength(1)
    expect(rows[0]).toMatchObject({ pid: 7, name: 'System Idle Process', command: 'System Idle Process', rssBytes: 8192 })

    expect(parseWinProcesses('not json')).toEqual([])
  })
})

describe('parseWinDisks', () => {
  it('parses drive letters with string and number sizes', () => {
    const text = JSON.stringify([
      { DeviceID: 'C:', Size: '500000000000', FreeSpace: '200000000000' },
      { DeviceID: 'D:', Size: 1000000000, FreeSpace: 500000000 }
    ])
    const disks = parseWinDisks(text)

    expect(disks).toEqual([
      { mount: 'C:\\', total: 500000000000, free: 200000000000, used: 300000000000 },
      { mount: 'D:\\', total: 1000000000, free: 500000000, used: 500000000 }
    ])
    expect(parseWinDisks('nope')).toEqual([])
  })
})

describe('parseWinBattery', () => {
  it('reads charge and whether it is on AC', () => {
    expect(parseWinBattery(JSON.stringify([{ EstimatedChargeRemaining: 42, BatteryStatus: 1 }]))).toEqual({ present: true, percent: 42, charging: false })
    expect(parseWinBattery(JSON.stringify([{ EstimatedChargeRemaining: '85', BatteryStatus: 2 }]))).toEqual({ present: true, percent: 85, charging: true })
  })

  it('reports no battery when there are no rows or the output is not JSON', () => {
    expect(parseWinBattery(JSON.stringify([]))).toEqual({ present: false })
    expect(parseWinBattery('garbage')).toEqual({ present: false })
  })
})

describe('parseNetWifi', () => {
  it('finds the Wi-Fi profile and its SSID among the connected profiles', () => {
    const text = JSON.stringify([
      { Name: 'Ethernet 1', InterfaceAlias: 'Ethernet' },
      { Name: 'Home 5G', InterfaceAlias: 'Wi-Fi' }
    ])
    expect(parseNetWifi(text)).toEqual({ connected: true, ssid: 'Home 5G' })
  })

  it('matches the Wi-Fi adapter name across spellings', () => {
    const text = JSON.stringify([{ Name: 'Guest', InterfaceAlias: 'WiFi' }])
    expect(parseNetWifi(text)).toEqual({ connected: true, ssid: 'Guest' })
  })

  it('returns null without a Wi-Fi profile or with non-JSON output', () => {
    expect(parseNetWifi(JSON.stringify([{ Name: 'Ethernet 1', InterfaceAlias: 'Ethernet' }]))).toBeNull()
    expect(parseNetWifi(JSON.stringify([]))).toBeNull()
    expect(parseNetWifi('')).toBeNull()
  })
})

describe('parseWinOs', () => {
  it('reads the OS caption and version', () => {
    const text = JSON.stringify([{ Caption: 'Microsoft Windows 11 Pro', Version: '10.0.26100' }])
    expect(parseWinOs(text)).toEqual({ osName: 'Microsoft Windows 11 Pro', osVersion: '10.0.26100' })
  })

  it('falls back on non-JSON output', () => {
    expect(parseWinOs('garbage')).toMatchObject({ osName: 'Windows' })
    expect(typeof parseWinOs('garbage').osVersion).toBe('string')
  })
})
