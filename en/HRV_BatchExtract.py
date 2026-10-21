# -*- coding: utf-8 -*-
"""
HRV BatchExtract - batch extraction of results from Kubios HRV Scientific Lite 4.3.0.
Automates the Kubios interface; every index is calculated by Kubios itself.

The program opens each .txt file in Kubios and applies:
  - Threshold (medium) artefact correction
  - analysis sample 00:00:00 to 00:05:00
It then copies the result tables (time domain, frequency domain FFT/AR and nonlinear)
and builds an Excel workbook. The values are Kubios's own (nothing is recalculated).

Do NOT use the mouse/keyboard during the run.
To stop: move the mouse quickly to the top-left corner of the screen,
or click "Stop" (the file in progress is finished first).
Processed files are saved in the subfolder "hrv_batchextract_results\\raw";
running again on the same folder skips them (automatic resume).
"""
import os, re, sys, time, glob, math, ctypes, threading, subprocess, traceback
from ctypes import wintypes

# ---------- DPI: use physical pixels for mouse and windows ----------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import pyautogui
import pyperclip
import pygetwindow as gw
import pandas as pd
import numpy as np

pyautogui.FAILSAFE = True      # mouse in the top-left corner = stop
pyautogui.PAUSE = 0.15

KUBIOS_EXE = r"C:\Program Files\Kubios\KubiosHRVScientific\application\launch_KubiosHRV.exe"
MAIN_TITLE = "Kubios HRV Scientific"
SAMPLE_LEN = "00:05:00"

# ---------- Coordinates (measured with Kubios MAXIMIZED, reference frame 1456 x 819) ----------
REF_W, REF_H = 1456.0, 819.0
P = {
    "menu_file":     (14, 32),
    "menu_open":     (50, 54),
    "beat_dropdown": (78, 263),
    "beat_medium":   (60, 343),
    "length_field":  (168, 393),
    "tab_time":      (234, 487),
    "tab_freq":      (338, 487),
    "tab_nonlin":    (442, 487),
    "tbl_time":      (470, 600),
    "tbl_freq":      (470, 612),
    "tbl_nonlin":    (333, 567),
}
CROP_CORR = (0, 276, 200, 292)      # text "Corrected: N beats (x %)"
CROP_BEAT = (4, 255, 152, 272)      # "Beat correction" menu text (e.g., "Threshold (medium)")
# "Text File Import" dialog: visible box (397,52)-(1060,762) in the reference frame
IMP = dict(L=397, T=52, W=663, H=710,
           units=(1028, 144), ms=(1012, 192), s=(1010, 205), importbtn=(688, 746))
# "Save Results" dialog: visible box (569,322)-(886,476); No button at (746,454)
SAVE = dict(L=569, T=322, W=317, H=154, no=(746, 454))

TEMPLATE_COLS = ['Code', 'SD1_percent_SNAP', 'SD2_percent_SNAS', 'PNS_Index', 'SNS_Index', 'Mean_RR', 'SDNN', 'Mean_HR', 'STD_HR', 'Min_HR', 'Max_HR', 'RMSSD', 'NNxx', 'pNNxx_percent', 'HRV_triangular_index', 'TINN', 'Stress_index', 'FFT_VLF_Peak_Hz', 'FFT_VLF_Power_ms2', 'FFT_VLF_Power_log', 'FFT_VLF_Power_percent', 'FFT_LF_Peak_Hz', 'FFT_LF_Power_ms2', 'FFT_LF_Power_log', 'FFT_LF_Power_percent', 'FFT_LF_Power_n.u.', 'FFT_HF_Peak_Hz', 'FFT_HF_Power_ms2', 'FFT_HF_Power_log', 'FFT_HF_Power_percent', 'FFT_HF_Power_n.u.', 'FFT_LF/HF_Power_ms2', 'AR_VLF_Peak_Hz', 'AR_VLF_Power_ms2', 'AR_VLF_Power_log', 'AR_VLF_Power_percent', 'AR_LF_Peak_Hz', 'AR_LF_Power_ms2', 'AR_LF_Power_log', 'AR_LF_Power_percent', 'AR_LF_Power_n.u.', 'AR_HF_Peak_Hz', 'AR_HF_Power_ms2', 'AR_HF_Power_log', 'AR_HF_Power_percent', 'AR_HF_Power_n.u.', 'AR_LF/HF_Power_ms2', 'Poincare_plot_SD1', 'Poincare_plot_SD2', 'Poincare_plot_SD2/SD1', 'ApEn', 'SampEn', 'DFA_alpha_1', 'DFA_alpha_2']

TMAP = {'PNS Index': 'PNS_Index', 'SNS Index': 'SNS_Index', 'Mean RR*': 'Mean_RR', 'SDNN': 'SDNN',
        'Mean HR*': 'Mean_HR', 'STD HR': 'STD_HR', 'Min HR': 'Min_HR', 'Max HR': 'Max_HR', 'RMSSD': 'RMSSD',
        'NNxx': 'NNxx', 'pNNxx': 'pNNxx_percent', 'HRV triangular index': 'HRV_triangular_index',
        'TINN': 'TINN', 'Stress index': 'Stress_index'}
NMAP = {'Poincare plot, SD1': 'Poincare_plot_SD1', 'Poincare plot, SD2': 'Poincare_plot_SD2',
        'Poincare plot, SD2/SD1': 'Poincare_plot_SD2/SD1', 'ApEn': 'ApEn', 'SampEn': 'SampEn',
        'DFA, alpha 1': 'DFA_alpha_1', 'DFA, alpha 2': 'DFA_alpha_2'}


class Stop(Exception):
    pass


# ======================= screen/window utilities =======================
def screen_size():
    u = ctypes.windll.user32
    return u.GetSystemMetrics(0), u.GetSystemMetrics(1)


def ref2scr(x, y):
    W, H = screen_size()
    return int(round(x * W / REF_W)), int(round(y * H / REF_H))


def visible_rect(win):
    """Visible window rectangle (without the invisible Windows 10/11 shadow)."""
    r = wintypes.RECT()
    try:
        ctypes.windll.dwmapi.DwmGetWindowAttribute(wintypes.HWND(win._hWnd), 9, ctypes.byref(r), ctypes.sizeof(r))
        return r.left, r.top, r.right - r.left, r.bottom - r.top
    except Exception:
        return win.left, win.top, win.width, win.height


def dlg_point(win, spec, key):
    """Converts a dialog point (reference frame) to screen coordinates, proportional to the actual size."""
    L, T, Wd, Hd = visible_rect(win)
    x, y = spec[key]
    return int(L + (x - spec['L']) / spec['W'] * Wd), int(T + (y - spec['T']) / spec['H'] * Hd)


def find_win(pred):
    for w in gw.getAllWindows():
        try:
            if w.title and pred(w.title):
                return w
        except Exception:
            pass
    return None


def main_win():
    return find_win(lambda t: t.startswith(MAIN_TITLE))


def activate(win):
    try:
        if win.isMinimized:
            win.restore()
        win.activate()
    except Exception:
        try:
            ctypes.windll.user32.SetForegroundWindow(win._hWnd)
        except Exception:
            pass
    time.sleep(0.4)


def focus_kubios():
    activate(main_win())
    x, y = ref2scr(700, 11)
    pyautogui.click(x, y)
    time.sleep(0.4)


def click_ref(name, clicks=1):
    x, y = ref2scr(*P[name])
    pyautogui.click(x, y, clicks=clicks, interval=0.08)


def hk(*keys):
    """Keyboard shortcut with pauses (Kubios/MATLAB misses shortcuts that are too fast)."""
    for k in keys:
        pyautogui.keyDown(k); time.sleep(0.08)
    for k in reversed(keys):
        time.sleep(0.05); pyautogui.keyUp(k)
    time.sleep(0.15)


def paste_text(txt):
    pyperclip.copy(txt)
    time.sleep(0.15)
    hk('ctrl', 'v')


def copy_table(tab, table, must_have, timeout=20):
    """Clicks the tab and the table, copies (ctrl+a, ctrl+c) and waits for the expected text."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        click_ref(tab)
        time.sleep(1.2)
        pyperclip.copy('')
        click_ref(table)
        hk('ctrl', 'a')
        hk('ctrl', 'c')
        time.sleep(0.5)
        txt = pyperclip.paste() or ''
        if must_have in txt:
            return txt.replace('\r', '')
        time.sleep(1.5)
    raise RuntimeError(f'Could not copy the table ({must_have})')


# ======================= optional OCR (native Windows OCR) =======================
def ocr_text(pil_img):
    try:
        import asyncio
        from winsdk.windows.media.ocr import OcrEngine
        from winsdk.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat, BitmapAlphaMode
        from winsdk.windows.storage.streams import DataWriter
        img = pil_img.convert('RGBA').resize((pil_img.width * 3, pil_img.height * 3))
        w, h = img.size
        data = bytearray(img.tobytes('raw', 'BGRA'))

        async def run():
            dw = DataWriter()
            dw.write_bytes(data)
            buf = dw.detach_buffer()
            bmp = SoftwareBitmap.create_copy_from_buffer(buf, BitmapPixelFormat.BGRA8, w, h, BitmapAlphaMode.PREMULTIPLIED)
            eng = OcrEngine.try_create_from_user_profile_languages()
            res = await eng.recognize_async(bmp)
            return res.text
        return asyncio.run(run())
    except Exception:
        return ''


def read_correction(raw_dir, cod):
    x1, y1 = ref2scr(*CROP_CORR[:2])
    x2, y2 = ref2scr(*CROP_CORR[2:])
    img = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
    os.makedirs(os.path.join(raw_dir, 'correction_screenshots'), exist_ok=True)
    img.save(os.path.join(raw_dir, 'correction_screenshots', f'{cod}.png'))
    t = ocr_text(img)
    m = re.search(r'(\d+)\s*beats?\s*\(\s*([\d.,]+)\s*%', t)
    if m:
        return int(m.group(1)), float(m.group(2).replace(',', '.')), t
    return None, None, t


# ======================= RR file reading =======================
def read_rr(path):
    txt = open(path, encoding='utf-8', errors='replace').read()
    vals = [float(v.replace(',', '.')) for v in re.findall(r'[-+]?\d+(?:[.,]\d+)?', txt)]
    vals = np.array([v for v in vals if v > 0])
    if len(vals) == 0:
        raise RuntimeError('file has no RR values')
    unit = 's' if np.median(vals) < 10 else 'ms'
    rr_ms = vals * 1000 if unit == 's' else vals
    return unit, rr_ms


def expected_mean_rr(rr_ms):
    """Typical RR within the first 300 s, used only to check the unit."""
    t = np.cumsum(rr_ms) / 1000.0
    sel = rr_ms[t <= 300.5]
    sel = sel[(sel > 250) & (sel < 2500)]     # ignore gross artefacts
    return float(np.median(sel)) if len(sel) else float(np.median(rr_ms))


# ======================= Kubios control =======================
def ensure_kubios(log, exe):
    w = main_win()
    if w is None:
        if not os.path.exists(exe):
            raise RuntimeError(f'Kubios not found at:\n{exe}')
        log('Opening Kubios (may take ~1 min)...')
        subprocess.Popen([exe])
        t0 = time.time()
        while time.time() - t0 < 240:
            time.sleep(2)
            w = main_win()
            if w:
                break
        if w is None:
            raise RuntimeError('Kubios did not open in time.')
        time.sleep(8)
    # close the "Kubios HRV" upgrade window, if present
    for _ in range(10):
        pop = find_win(lambda t: t.strip() == 'Kubios HRV')
        if not pop:
            break
        try:
            pop.close()
        except Exception:
            pass
        time.sleep(1)
    w = main_win()
    activate(w)
    try:
        if not w.isMaximized:
            w.maximize()
            time.sleep(1.5)
    except Exception:
        pass
    return w


def wait_loaded_or_import(fname, timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        d = find_win(lambda t: t.startswith('Text File Import'))
        if d:
            return 'import', d
        w = main_win()
        if w and w.title.replace('/', '\\').lower().endswith(fname.lower()):
            return 'loaded', w
        time.sleep(0.7)
    return 'timeout', None


def open_file(path, unit, log, retry=False):
    fname = os.path.basename(path)
    w = main_win()
    activate(w)
    x, y = ref2scr(700, 11)          # click the Kubios title bar to make sure it has focus
    pyautogui.click(x, y)
    time.sleep(0.5)
    click_ref('menu_file')           # File > Open
    time.sleep(1.2)
    click_ref('menu_open')
    # If a file was already open, Kubios asks "Save Results?" (dialog drawn inside the main window).
    # Tab moves the focus from "Yes" to "No"; Enter confirms.
    o = None
    for attempt in range(3):
        t0 = time.time()
        while time.time() - t0 < 4:
            o = find_win(lambda t: t.startswith('Open Data File'))
            if o:
                break
            time.sleep(0.4)
        if o:
            break
        s = find_win(lambda t: t.startswith('Save Results'))
        if s:
            activate(s)
            pyautogui.click(*dlg_point(s, SAVE, 'no'))
        else:
            pyautogui.press('tab'); time.sleep(0.3); pyautogui.press('enter')
        time.sleep(1.5)
    if o is None:
        titles = [t for t in gw.getAllTitles() if t.strip()]
        raise RuntimeError('The "Open Data File" window did not appear. Open windows: ' + ' | '.join(titles[:15]))
    activate(o)
    time.sleep(0.5)
    paste_text(path)
    pyautogui.press('enter')
    time.sleep(7 if retry else 2)

    state, win = wait_loaded_or_import(fname)
    if state == 'import':
        activate(win)
        time.sleep(0.6)
        pyautogui.click(*dlg_point(win, IMP, 'units'))
        time.sleep(0.8)
        pyautogui.click(*dlg_point(win, IMP, unit))
        time.sleep(0.8)
        pyautogui.click(*dlg_point(win, IMP, 'importbtn'))
        time.sleep(2)
        state, win = wait_loaded_or_import(fname)
        if state != 'loaded':
            raise RuntimeError('The file did not finish loading after import.')
    elif state != 'loaded':
        raise RuntimeError('The file did not load (timed out).')
    time.sleep(3)
    focus_kubios()


def screen_text(box):
    x1, y1 = ref2scr(*box[:2])
    x2, y2 = ref2scr(*box[2:])
    return ocr_text(pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1)))


def close_popup():
    pop = find_win(lambda t: t.strip() == 'Kubios HRV')
    if pop:
        try:
            pop.close()
        except Exception:
            pass
        time.sleep(1)
        focus_kubios()


def ensure_medium(log):
    """Ensures 'Threshold (medium)' on EVERY file and verifies it by reading the menu text on screen."""
    for attempt in range(5):
        txt = screen_text(CROP_BEAT)
        low = txt.lower().replace(' ', '')
        if 'medium' in low:
            return txt
        if attempt == 0 and not txt:
            log('  (warning: screen reading unavailable; applying the filter without verification)')
        if attempt % 2 == 0:
            click_ref('beat_dropdown')
            time.sleep(1.2)
            click_ref('beat_medium')
        else:                               # keyboard alternative
            click_ref('beat_dropdown')
            time.sleep(1.0)
            pyautogui.press('home'); time.sleep(0.2)
            for _ in range(3):
                pyautogui.press('down'); time.sleep(0.2)
            pyautogui.press('enter')
        time.sleep(4)
        if not txt and attempt == 0:
            txt2 = screen_text(CROP_BEAT)
            if not txt2:                   # no OCR: cannot verify, continue with the click already made
                return '(not verified)'
    raise RuntimeError('Could not apply Threshold (medium) correction. Text read: ' + repr(txt))


def set_length():
    click_ref('length_field', clicks=3)
    time.sleep(0.3)
    paste_text(SAMPLE_LEN)
    pyautogui.press('enter')
    time.sleep(4)


# ======================= parse =======================
def num(s):
    s = s.strip()
    try:
        return float(s)
    except Exception:
        return np.nan


def parse_all(cod, t_time, t_freq, t_nl):
    r = {'Code': cod}
    for ln in t_time.splitlines():
        p = ln.split('\t')
        if p[0] in TMAP:
            r[TMAP[p[0]]] = num(p[1])
    sec = None
    for ln in t_freq.splitlines():
        p = ln.split('\t')
        if p[0].startswith('FFT Results'):
            sec = 'FFT'; continue
        if p[0].startswith('AR Results'):
            sec = 'AR'; continue
        if sec is None:
            continue
        v = [num(x) for x in p[1:]] + [np.nan] * 5
        k, S = p[0], sec
        if k == 'Peak (Hz)':
            r[S+'_VLF_Peak_Hz'], r[S+'_LF_Peak_Hz'], r[S+'_HF_Peak_Hz'] = v[:3]
        elif k == 'Power (ms2)':
            r[S+'_VLF_Power_ms2'], r[S+'_LF_Power_ms2'], r[S+'_HF_Power_ms2'], r[S+'_LF/HF_Power_ms2'] = v[:4]
        elif k == 'Power (log)':
            r[S+'_VLF_Power_log'], r[S+'_LF_Power_log'], r[S+'_HF_Power_log'] = v[:3]
        elif k == 'Power (%)':
            r[S+'_VLF_Power_percent'], r[S+'_LF_Power_percent'], r[S+'_HF_Power_percent'] = v[:3]
        elif k == 'Power (n.u.)':
            r[S+'_LF_Power_n.u.'], r[S+'_HF_Power_n.u.'] = v[1:3]
    t_nl = t_nl.replace('\t ', ', ')     # variant "Poincare plot\t SD1"
    for ln in t_nl.splitlines():
        p = ln.split('\t')
        if p[0] in NMAP:
            r[NMAP[p[0]]] = num(p[1])
    return r


# ======================= processing of one file =======================
def process_file(path, raw_dir, log, first):
    cod = re.sub(r'_5min$', '', os.path.splitext(os.path.basename(path))[0])
    unit, rr = read_rr(path)
    exp_rr = expected_mean_rr(rr)
    tried = []
    for u in [unit, 'ms' if unit == 's' else 's']:
        tried.append(u)
        close_popup()
        open_file(path, u, log, retry=len(tried) > 1)
        close_popup()
        beat_filter = ensure_medium(log)
        set_length()
        t_time = copy_table('tab_time', 'tbl_time', 'RMSSD')
        t_freq = copy_table('tab_freq', 'tbl_freq', 'AR Results')
        t_nl = copy_table('tab_nonlin', 'tbl_nonlin', 'alpha 2')
        r = parse_all(cod, t_time, t_freq, t_nl)
        mrr = r.get('Mean_RR', np.nan)
        ok_unit = np.isfinite(mrr) and abs(mrr - exp_rr) / exp_rr < 0.15
        if ok_unit:
            break
        log(f'  {cod}: Kubios Mean RR ({mrr:.1f}) does not match the file ({exp_rr:.1f}); trying the other unit...')
    n_corr, p_corr, ocr_raw = read_correction(raw_dir, cod)
    sd1, rmssd = r.get('Poincare_plot_SD1', np.nan), r.get('RMSSD', np.nan)
    ok_sd1 = np.isfinite(sd1) and np.isfinite(rmssd) and abs(sd1 - rmssd / math.sqrt(2)) < 0.02 * sd1 + 0.05
    check = {'Code': cod, 'File_unit': unit, 'Unit_used': tried[-1],
             'MeanRR_file_approx': round(exp_rr, 1), 'MeanRR_Kubios': mrr,
             'Unit_OK': bool(ok_unit), 'SD1_vs_RMSSD_OK': bool(ok_sd1),
             'Kubios_filter': beat_filter, 'Corrected_beats': n_corr, 'Corrected_beats_percent': p_corr}
    with open(os.path.join(raw_dir, 'raw', f'{cod}.txt'), 'w', encoding='utf-8') as f:
        f.write('#FILE\t' + path + '\n')
        f.write('#CHECK\t' + repr(check) + '\n')
        f.write('#OCR\t' + (ocr_raw or '').replace('\n', ' ') + '\n')
        f.write('#TIME\n' + t_time + '\n#FREQ\n' + t_freq + '\n#NONLIN\n' + t_nl + '\n')
    return r, check


def load_raw(fn):
    txt = open(fn, encoding='utf-8').read()
    check = eval(re.search(r'^#CHECK\t(.*)$', txt, re.M).group(1), {'nan': np.nan})
    t_time = txt.split('#TIME\n')[1].split('\n#FREQ\n')[0]
    t_freq = txt.split('#FREQ\n')[1].split('\n#NONLIN\n')[0]
    t_nl = txt.split('#NONLIN\n')[1]
    return parse_all(check['Code'], t_time, t_freq, t_nl), check


# ======================= Excel =======================
def sort_key(c):
    return (0, int(c), c) if str(c).isdigit() else (1, 0, str(c))


def build_excel(folder, raw_dir):
    rows, checks = [], []
    for fn in glob.glob(os.path.join(raw_dir, 'raw', '*.txt')):
        r, c = load_raw(fn)
        rows.append(r); checks.append(c)
    if not rows:
        return None
    df = pd.DataFrame(rows)
    sd1, sd2 = df['Poincare_plot_SD1'], df['Poincare_plot_SD2']
    df['SD1_percent_SNAP'] = (100 * sd1 / (sd1 + sd2)).round(0)
    df['SD2_percent_SNAS'] = (100 * sd2 / (sd1 + sd2)).round(0)
    for c in TEMPLATE_COLS:
        if c not in df.columns:
            df[c] = np.nan
    df = df.iloc[sorted(range(len(df)), key=lambda i: sort_key(df.Code.iloc[i]))][TEMPLATE_COLS]
    ck = pd.DataFrame(checks)
    ck = ck.iloc[sorted(range(len(ck)), key=lambda i: sort_key(ck.Code.iloc[i]))]
    met = pd.DataFrame({'Item': ['Software', 'Artefact correction', 'Sample', 'Domains', 'SD1%/SD2%',
                                 'Verification', 'Corrected beats'],
                        'Description': ['Indices calculated in Kubios HRV Scientific Lite 4.3.0; result extraction automated by HRV BatchExtract (Python script that copies the result tables from the Kubios interface).',
                                      'Threshold-based, medium, applied and verified on every file (Kubios_filter column of the Verification sheet).',
                                      'Start 00:00:00, length 00:05:00 (shorter recordings: Kubios uses the full available length).',
                                      'Time, frequency (FFT and AR) and nonlinear (Poincare, ApEn, SampEn, DFA).',
                                      'SD1% = 100*SD1/(SD1+SD2); SD2% = 100*SD2/(SD1+SD2), rounded.',
                                      'Verification sheet: Unit_OK compares the Kubios Mean RR with that of the file; SD1_vs_RMSSD_OK checks SD1 ~ RMSSD/sqrt(2).',
                                      'Read from the screen by Windows OCR; the image of each file is saved in hrv_batchextract_results\\correction_screenshots. If the column is empty, check the image.']})
    out = os.path.join(raw_dir, 'HRV_BatchExtract.xlsx')
    try:
        with pd.ExcelWriter(out) as w:
            df.to_excel(w, sheet_name='HRV', index=False)
            ck.to_excel(w, sheet_name='Verification', index=False)
            met.to_excel(w, sheet_name='Method', index=False)
    except PermissionError:
        out = os.path.join(raw_dir, f'HRV_BatchExtract_{time.strftime("%Y%m%d_%H%M%S")}.xlsx')
        with pd.ExcelWriter(out) as w:
            df.to_excel(w, sheet_name='HRV', index=False)
            ck.to_excel(w, sheet_name='Verification', index=False)
            met.to_excel(w, sheet_name='Method', index=False)
    return out, ck


# ======================= user interface =======================
class App:
    def __init__(self, root):
        self.root = root
        root.title('HRV BatchExtract')
        root.geometry('640x460')
        self.stop = threading.Event()
        self.folder = tk.StringVar()
        self.exe = tk.StringVar(value=KUBIOS_EXE)
        f = tk.Frame(root, padx=10, pady=8); f.pack(fill='x')
        tk.Label(f, text='Folder with the .txt files:').grid(row=0, column=0, sticky='w')
        tk.Entry(f, textvariable=self.folder, width=62).grid(row=1, column=0, sticky='we')
        tk.Button(f, text='Browse...', command=self.pick).grid(row=1, column=1, padx=5)
        tk.Label(f, text='Kubios:').grid(row=2, column=0, sticky='w', pady=(6, 0))
        tk.Entry(f, textvariable=self.exe, width=62).grid(row=3, column=0, sticky='we')
        b = tk.Frame(root, pady=6); b.pack()
        self.btn = tk.Button(b, text='  RUN  ', font=('Segoe UI', 14, 'bold'), bg='#2e7d32', fg='white', command=self.start)
        self.btn.pack(side='left', padx=8)
        tk.Button(b, text='Stop', command=self.stop.set).pack(side='left', padx=8)
        tk.Button(b, text='Build Excel only', command=self.only_excel).pack(side='left', padx=8)
        self.logbox = scrolledtext.ScrolledText(root, height=16, font=('Consolas', 9))
        self.logbox.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        self.log('Choose the folder and click RUN. Do not use the mouse/keyboard during the run.')
        self.log('To stop immediately: move the mouse to the top-left corner of the screen.')

    def log(self, msg):
        def _():
            self.logbox.insert('end', msg + '\n'); self.logbox.see('end')
        self.root.after(0, _)

    def pick(self):
        d = filedialog.askdirectory()
        if d:
            self.folder.set(os.path.normpath(d))

    def only_excel(self):
        folder = self.folder.get().strip()
        raw_dir = os.path.join(folder, 'hrv_batchextract_results')
        res = build_excel(folder, raw_dir) if os.path.isdir(raw_dir) else None
        if res:
            self.log(f'Workbook: {res[0]}')
        else:
            self.log('Nothing processed in this folder yet.')

    def start(self):
        folder = self.folder.get().strip()
        if not os.path.isdir(folder):
            messagebox.showerror('HRV BatchExtract', 'Choose a valid folder.'); return
        self.stop.clear()
        self.btn.config(state='disabled')
        threading.Thread(target=self.run, args=(folder,), daemon=True).start()

    def run(self, folder):
        raw_dir = os.path.join(folder, 'hrv_batchextract_results')
        os.makedirs(os.path.join(raw_dir, 'raw'), exist_ok=True)
        files = sorted([f for f in glob.glob(os.path.join(folder, '*.txt'))],
                       key=lambda p: sort_key(re.sub(r'_5min$', '', os.path.splitext(os.path.basename(p))[0])))
        done = {os.path.splitext(f)[0] for f in os.listdir(os.path.join(raw_dir, 'raw'))}
        todo = [p for p in files if re.sub(r'_5min$', '', os.path.splitext(os.path.basename(p))[0]) not in done]
        self.log(f'{len(files)} .txt files; {len(files) - len(todo)} already done; {len(todo)} to process.')
        errors = []
        try:
            self.root.after(0, self.root.iconify)
            time.sleep(1)
            ensure_kubios(self.log, self.exe.get())
            first = [True]
            t0 = time.time()
            for i, p in enumerate(todo, 1):
                if self.stop.is_set():
                    self.log('Stopped by the user.'); break
                name = os.path.basename(p)
                try:
                    r, c = process_file(p, raw_dir, self.log, first)
                    flag = '' if (c['Unit_OK'] and c['SD1_vs_RMSSD_OK']) else '  <-- CHECK'
                    el = (time.time() - t0) / i
                    self.log(f'[{i}/{len(todo)}] {name}: RMSSD {r.get("RMSSD")}  SD1 {r.get("Poincare_plot_SD1")}  '
                             f'corr {c["Corrected_beats"]}{flag}  (~{el * (len(todo) - i) / 60:.0f} min remaining)')
                except pyautogui.FailSafeException:
                    self.log('Stopped (mouse in the screen corner).'); break
                except Exception as e:
                    errors.append(name)
                    self.log(f'[{i}/{len(todo)}] ERROR in {name}: {e}')
                    try:
                        for t in ('Text File Import', 'Open Data File', 'Save Results'):
                            d = find_win(lambda x, t=t: x.startswith(t))
                            if d:
                                activate(d); pyautogui.press('esc'); time.sleep(1)
                    except Exception:
                        pass
                    if len(errors) >= 5 and len(errors) == i:
                        self.log('5 consecutive errors at the start: stopping. Check that Kubios is maximized and visible.')
                        break
        except pyautogui.FailSafeException:
            self.log('Stopped (mouse in the screen corner).')
        except Exception as e:
            self.log('ERROR: ' + str(e)); self.log(traceback.format_exc())
        finally:
            try:
                res = build_excel(folder, raw_dir)
                if res:
                    out, ck = res
                    bad = ck[~(ck.Unit_OK & ck.SD1_vs_RMSSD_OK)].Code.tolist()
                    self.log(f'Workbook created: {out}')
                    if bad:
                        self.log('Check: ' + ', '.join(map(str, bad)))
                if errors:
                    self.log('With errors (run again to retry only these): ' + ', '.join(errors))
            except Exception as e:
                self.log('Error building Excel: ' + str(e))
            self.root.after(0, self.root.deiconify)
            self.root.after(0, lambda: self.btn.config(state='normal'))


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
