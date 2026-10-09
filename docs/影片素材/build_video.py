# -*- coding: utf-8 -*-
"""教學影片合成：每句旁白 = 一段（畫面＋字幕＋語音），最後串接。
規格沿用 tra-art 教學影片：女聲 zh-TW-HsiaoChenNeural +5%、1920x1080 30fps、底部 150px 深藍字幕條。
用法：先跑 render_cards.py 產生字卡，再跑  python docs\\影片素材\\build_video.py [完整版|操作篇]（預設完整版）
"""
import asyncio, hashlib, json, os, re, subprocess, sys
from PIL import Image, ImageDraw, ImageFont
import edge_tts

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.dirname(HERE)
OUT = r"D:\320code\material-price-payment-教學影片"
WORK = os.path.join(os.environ.get("TEMP", "."), "mpp_video_work")
NAME = "估驗計價核算工具_操作教學"
VOICE, RATE = "zh-TW-HsiaoChenNeural", "+5%"
W, H, CAP_H = 1920, 1080, 150
FONT = r"C:\Windows\Fonts\msjhbd.ttc"
NAVY, ORANGE, BG = (15, 42, 71), (226, 97, 29), (226, 232, 240)

def card(n): return os.path.join(HERE, "字卡", n + ".png")
def shot(n): return os.path.join(HERE, n)
def img(n): return os.path.join(DOCS, "img", n)
IMG = {
    "github": shot("01-GitHub首頁.png"), "zip": shot("02-GitHub-Download-ZIP.png"),
    "m0": shot("03-手冊-第0步安裝.png"), "m2": shot("04-手冊-第2步複製這段話.png"),
    "m3": shot("05-手冊-第3步看結論.png"), "mans": shot("06-手冊-練習對答案.png"),
    "m5": shot("07-手冊-第5步檢核清單.png"),
    "rA": shot("08-報告-A-跨期銜接比對.png"), "rB": shot("09-報告-B-前期加本期不等於累計.png"),
    "rC": shot("10-報告-C-畫面顯示比對.png"), "rD": shot("11-報告-D-略過未檢查.png"),
    "rTot": shot("12-報告-計價總表結論.png"), "rDet": shot("13-報告-請款明細表結論.png"),
    "x01": img("01-請款明細表.png"), "x02": img("02-A-第一期累計.png"), "x03": img("03-A-第二期前期累計.png"),
    "x04": img("04-B-固定值誤植.png"), "x05": img("05-C-第一期顯示.png"), "x06": img("06-C-第二期顯示.png"),
}
ROW11 = (50, 327, 1527, 355)  # 01-請款明細表.png 的「物價調整金額」列

# 唸法修正（只影響語音，字幕照原文）；長的先換
SAY = [("0.367188", "零點三六七一八八"), ("0.37", "零點三七"), ("0.3", "零點三"), ("12,500", "一萬兩千五百"),
       ("12,000", "一萬兩千"), ("30%", "百分之三十"), ("J7", "J 7"), ("J8", "J 8"), ("320code", "三二零 code")]

# (畫面, 重點框[], 旁白)；畫面 cXX／r0X = 字卡（r0X 是取代操作錄影的示意字卡）
SCENES = [
    # ── 第一段　能做什麼 ──
    ("c01", [], "這是估驗計價逐項核算工具的操作教學。今天分三段：先講它能做什麼，再講怎麼用，最後複習今天學到什麼。"),
    ("x01", [ROW11], "每一期估驗計價送進來，我們都要逐列對。單價乘數量對不對，前期加本期是不是等於累計。這一期的前期累計，跟上一期的累計至本期，接不接得起來。"),
    ("c02", [], "表格一長，最容易漏看的就是這幾件事。尤其是兩期的銜接數字，前一期改了，本期常常忘記同步。"),
    ("c03", [], "這個工具，就是把前一期跟本期兩份 Excel 交給 Claude，說一句話，它就幫你逐列核對，對不起來的地方直接指出來。所有數字都是程式從 Excel 讀出來算的，AI 只負責把結果翻成白話。要不要改、怎麼改，最後還是你決定。"),
    ("c04_hl1", [], "它一共核五件事。第一，每一列的單價乘本期數量，是不是等於複價。"),
    ("c04_hl2", [], "第二，同一份檔案裡，前期加本期是不是等於累計。"),
    ("c04_hl3", [], "第三，前一期的累計至本期，跟本期的前期累計一不一樣。這是最常忘記同步的地方。"),
    ("c04_hl4", [], "第四，小計列是不是真的等於底下子項加起來。"),
    ("c04_hl5", [], "第五，兩期同一個數字，畫面上顯示出來長得一不一樣。數字一樣，但一期顯示兩位小數、一期顯示六位，審查的人看紙本會以為對不起來，這個它也會抓。"),
    ("c05", [], "等一下用練習檔示範，你會看到它抓出四種最常見的狀況：前一期改了本期沒跟著改，公式被打成固定數字，數字一樣但顯示格式不一樣，還有它沒辦法核的地方，會老實標出來。"),
    # ── 第二段　怎麼用 ──
    ("c06", [], "接下來講怎麼用。只有三個動作：準備兩份檔案、說一句話、看結論。不過第一次用，要先安裝。"),
    ("github", [], "先到這個 GitHub 網址。"),
    ("zip", [], "按綠色的 Code，選 Download ZIP，解壓縮到 D 槽 320code 底下的 material-price-payment 資料夾。"),
    ("m0", [], "打開 PowerShell，把手冊上這兩行貼上去，按 Enter。"),
    ("r01", [], "看到安裝完成，再把 Claude Code 關掉重開，工具就裝好了。"),
    ("r02", [], "動作一，準備兩份檔案。我們先用練習檔示範。這兩份是假資料，裡面故意埋了幾種常見錯誤，可以放心亂試。"),
    ("x01", [], "表格長這樣：左邊是項次和工程項目，右邊依序是原契約、前期累計、本期完成、累計至本期。工具會自己找這幾個標題在哪一欄，不同機關的範本大多可以直接用。"),
    ("c07", [], "準備檔案注意三件事。檔案要先關掉；檔名開頭建議放日期，一眼知道誰先誰後；請款明細表跟計價總表，兩張都要核。"),
    ("m2", [], "動作二，說一句話。手冊上有一段現成的話，按複製，貼到 Claude Code 送出。真正用的時候，把路徑換成你自己的檔案。"),
    ("c08", [], "最重要的一句：一定要講清楚哪一份是前一期。順序反了，跨期比對會整個錯。"),
    ("r03", [], "送出之後，它會自己跑核算程式，兩張分頁各跑一次，報告存在前一期檔案旁邊。"),
    ("rDet", [], "動作三，看結論。先看最後那一句。請款明細表的結論是：共發現四筆異常。"),
    ("rTot", [], "計價總表是共發現八筆異常，另外有兩項檢查沒跑到。"),
    ("m3", [], "結論有四種。全數相符就是都對；發現幾筆異常就逐筆看；出現略過的警告，代表有一項根本沒核到；無法核算，通常是路徑或分頁名稱打錯。"),
    ("mans", [], "練習檔應該得到：請款明細表四筆，計價總表八筆加兩項略過。跟這個一樣，就代表你裝好了。"),
    ("c09", [], "計價總表報八筆，看起來很多，其實只有三個原因，再加一項沒核到。我們一個一個看。"),
    ("rA", [], "狀況 A：前一期改了，本期沒跟著改。跨期比對抓到物價調整金額，前一期累計 12,000，本期的前期累計卻是 12,500，總計也差五百。"),
    ("x02", [], "前一期這裡是修正後的 12,000。"),
    ("x03", [], "本期還留著修正前的 12,500，總計跟著差。"),
    ("x03", [], "兩筆差額一樣都是五百，所以是同一件事，不是兩個問題。處理方式：先確認前一期的數字才是對的，再把本期的前期累計改過來，兩張表都要改。"),
    ("rB", [], "狀況 B：公式被打成固定數字。同一份檔案自己加起來就不對：交通工程費，金額對，數量卻不對。"),
    ("r04", [], "回到 Excel，按 Ctrl 加數字 1 左邊那個鍵，整張表會顯示公式。"),
    ("x04", [], "上一列 J7 是公式，J8 卻是固定數字 30。應該是 0.3，有人把 30% 打成 30。"),
    ("x04", [], "把這格改回公式就好。這一格也連帶讓跨期比對、畫面比對各報一次，改好會一起消失。"),
    ("rC", [], "狀況 C：數字一樣，畫面長得不一樣。報告會直接標出格式不同、數字相同。"),
    ("pairC", [], "前一期顯示 0.37，本期顯示 0.367188，背後其實是同一個數。"),
    ("r05", [], "數字不用改，按 Ctrl 加 1，把兩期同一欄的小數位數設成一樣就好。"),
    ("rD", [], "狀況 D：計價總表通常沒有單價欄，單價乘數量這一項沒辦法核，報告會標略過。"),
    ("c10", [], "略過不是核過沒問題，是沒核到。這一項請人工抽查，或以請款明細表的結果為準。另外，百分比費用列、變更設計欄、手動加總的小計，它也不核，請人工驗算。"),
    ("c11", [], "Excel 改完存檔，回去再跑一次這個技能。重點不是跑完一次就好，而是問題到底有沒有解決。還有異常，就再修正、再跑一次，直到沒有異常為止；略過的項目，用人工驗算補上，事情才算結束。"),
    ("m5", [], "手冊最後有一張檢核清單，可以印出來，每次核完照著勾。"),
    # ── 第三段　今天學到什麼 ──
    ("c12", [], "最後複習一下，今天學到什麼。"),
    ("c12_hl1", [], "第一，這個工具幫你核五件事：乘法、前期加本期、跨期銜接、小計，還有畫面顯示。"),
    ("c12_hl2", [], "第二，用法就三個動作：準備兩份檔案、說一句話、看結論。記得前一期一定放前面。"),
    ("c12_hl3", [], "第三，看到異常先分辨是哪一種：前一期改了本期沒改、公式被打成數字、只是格式不同，或是根本沒核到。差額一樣的多筆，通常是同一個原因。"),
    ("c12_hl4", [], "第四，有錯就修正後再跑一次，直到沒有異常；略過的項目用人工驗算，才算完成。"),
    ("c13", [], "手冊在 docs 資料夾裡的快速上手。建議第一次先拿練習檔跑一遍、對過答案，再用在真的案子上。謝謝收看。"),
]

# 實際操作篇（Claude Code 桌面 App，畫面為示意字卡 o01～o13）
SCENES_OP = [
    ("o01", [], "這是估驗計價核算工具的實際操作篇。從打開 Claude，到拿到核算結果，帶你一步一步跑一次。"),
    ("o02", [], "一共五步：開啟 Claude Code、開啟技能、貼上指令送出、看結果，最後是修正後再跑一次。"),
    ("o03", [], "第一步，打開 Claude 桌面程式，點上方的 Code 分頁。"),
    ("o04", [], "按新工作階段，資料夾選 D 槽 320code 底下的 material-price-payment。練習檔就在這個資料夾裡，最方便。"),
    ("o05", [], "第二步，開啟技能。在輸入框打一個斜線，清單裡會出現 material-price-payment，點它就開啟了。"),
    ("o06", [], "也可以不打斜線，直接說核對估驗計價、跨期比對這類關鍵字，技能會自己啟動。兩種方法都可以。"),
    ("o07", [], "第三步，把手冊上那段話貼上去，按 Enter 送出。記得前一期在前、本期在後。"),
    ("o08", [], "如果畫面跳出詢問，要不要讓它執行核算程式，按允許。它只是讀你的 Excel 來核算，不會改到你的 Excel。"),
    ("o09", [], "接著它會自己跑，請款明細表和計價總表各跑一次，稍等一下就好。"),
    ("o10", [], "第四步，看結果。先看結論那一句：請款明細表四筆異常，計價總表八筆異常，另外有兩項略過。"),
    ("o11", [], "核算報告存在前一期檔案旁邊，每張分頁一份，想看細節可以打開來看。"),
    ("o12", [], "看不懂的地方直接問，例如：這筆異常是什麼原因、要怎麼改。它會照核算結果，用白話跟你解釋。"),
    ("o14", [], "第五步，也是最重要的一步：修正後再跑一次。重點不是跑完一次就好，而是問題到底有沒有解決。有異常，就到 Excel 修正、存檔，再跑一次這個技能。還有異常就重複，直到沒有異常為止。"),
    ("o15", [], "改好之後，直接跟它說：我改好了，再核一次。它會用同樣的兩份檔案重新核算，告訴你問題是不是都解決了。"),
    ("o16", [], "什麼時候才算結束？結論沒有異常，而且略過、不核的項目，都已經人工驗算過，事情才算結束。只跑一次就交差，不算完成。"),
    ("o13", [], "複習一下：開啟 Claude Code、開啟技能、貼上送出、看結果，有錯就修正後再跑，直到沒有異常為止。異常要怎麼判斷、怎麼改，請看完整版的教學影片。謝謝收看。"),
]
VIDEOS = {"完整版": (NAME, SCENES), "操作篇": ("估驗計價核算工具_實際操作篇", SCENES_OP)}

def split_sentences(t):
    parts = [p for p in re.split(r"(?<=[。；：])", t) if p.strip()]
    out = []
    for p in parts:  # 冒號短句併入下一句
        if out and out[-1].endswith("："): out[-1] += p
        else: out.append(p)
    return out

def font(sz): return ImageFont.truetype(FONT, sz)

def wrap(draw, text, f, maxw):
    lines, cur = [], ""
    for ch in text:
        if draw.textlength(cur + ch, font=f) > maxw:
            lines.append(cur); cur = ch
        else: cur += ch
    if cur: lines.append(cur)
    return lines

def fit(im, box_w, box_h):
    s = min(box_w / im.width, box_h / im.height)
    return im.resize((round(im.width * s), round(im.height * s)), Image.LANCZOS), s

def base_frame(key, boxes):
    if os.path.exists(card(key)):  # 字卡等比縮放置中，左右用邊緣色補滿（不變形）
        im, _ = fit(Image.open(card(key)).convert("RGB"), W, H - CAP_H)
        canvas = Image.new("RGB", (W, H - CAP_H))
        x0 = (W - im.width) // 2
        canvas.paste(im.crop((0, 0, 1, im.height)).resize((x0 + 1, im.height)), (0, 0))
        canvas.paste(im.crop((im.width - 1, 0, im.width, im.height)).resize((x0 + 1, im.height)), (W - x0 - 1, 0))
        canvas.paste(im, (x0, 0))
        return canvas
    canvas = Image.new("RGB", (W, H - CAP_H), BG)
    d = ImageDraw.Draw(canvas)
    if key == "pairC":  # 前一期／本期上下對照
        a, b = (Image.open(IMG[k]).convert("RGB") for k in ("x05", "x06"))
        ia, s = fit(a, W - 320, 330); ib, _ = fit(b, W - 320, 330)
        x0 = (W - ia.width) // 2
        for i, (lab, im) in enumerate((("前一期：J7 顯示 0.37", ia), ("本期：F7 顯示 0.367188", ib))):
            y = 50 + i * (im.height + 110)
            d.text((x0, y), lab, fill=NAVY, font=font(44))
            canvas.paste(im, (x0, y + 62))
        return canvas
    im0 = Image.open(IMG[key]).convert("RGB")
    im, s = fit(im0, W - 80, H - CAP_H - 40)
    x0, y0 = (W - im.width) // 2, (H - CAP_H - im.height) // 2
    canvas.paste(Image.new("RGB", (im.width + 12, im.height + 12), (200, 208, 220)), (x0 - 6, y0 - 2))
    canvas.paste(im, (x0, y0))
    for b in boxes:
        d.rounded_rectangle((x0 + b[0] * s - 6, y0 + b[1] * s - 6, x0 + b[2] * s + 6, y0 + b[3] * s + 6), 10, outline=ORANGE, width=6)
    return canvas

def compose(base, caption):
    fr = Image.new("RGB", (W, H), NAVY); fr.paste(base, (0, 0))
    d = ImageDraw.Draw(fr); f = font(46)
    lines = wrap(d, caption, f, W - 200)
    lh = 60; y = H - CAP_H + (CAP_H - lh * len(lines)) // 2
    for ln in lines:
        d.text(((W - d.textlength(ln, font=f)) / 2, y), ln, fill=(255, 255, 255), font=f); y += lh
    return fr

def dur(p):
    r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "json", p], capture_output=True, text=True)
    return float(json.loads(r.stdout)["format"]["duration"])

def ts(t):
    h, m = int(t // 3600), int(t % 3600 // 60); s = t % 60
    return f"{h:02d}:{m:02d}:{int(s):02d},{int(round((s - int(s)) * 1000)):03d}"

def say(t):
    for a, b in SAY: t = t.replace(a, b)
    return t

async def main(name, scenes):
    os.makedirs(WORK, exist_ok=True); os.makedirs(OUT, exist_ok=True)
    segs, srt, t, n = [], [], 0.0, 0
    for si, (key, boxes, text) in enumerate(scenes):
        base = base_frame(key, boxes)
        for sent in split_sentences(text):
            n += 1
            spoken = say(sent)
            mp3 = os.path.join(WORK, "a_" + hashlib.md5((VOICE + RATE + spoken).encode()).hexdigest()[:12] + ".mp3")
            png = os.path.join(WORK, f"f{n:03d}.png"); mp4 = os.path.join(WORK, f"s{n:03d}.mp4")
            if not os.path.exists(mp3):  # 旁白沒改的句子不重新合成語音
                await edge_tts.Communicate(spoken, VOICE, rate=RATE).save(mp3)
            cap = sent.rstrip("。；")
            compose(base, cap).save(png)
            d = dur(mp3) + 0.35
            subprocess.run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", "30", "-i", png, "-i", mp3,
                            "-af", "apad", "-t", f"{d:.3f}", "-c:v", "libx264", "-tune", "stillimage", "-pix_fmt", "yuv420p",
                            "-r", "30", "-c:a", "aac", "-b:a", "160k", "-ar", "48000", "-ac", "2", mp4], check=True)
            segs.append(mp4)
            srt.append(f"{n}\n{ts(t)} --> {ts(t + d - 0.2)}\n{cap}\n")
            t += d
        print(f"scene {si + 1}/{len(scenes)} done, t={t:.1f}s", flush=True)
    lst = os.path.join(WORK, "list.txt")
    with open(lst, "w", encoding="utf-8") as fh:
        fh.writelines(f"file '{p}'\n" for p in segs)
    out = os.path.join(OUT, name + ".mp4")
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-f", "concat", "-safe", "0", "-i", lst, "-c", "copy", "-movflags", "+faststart", out], check=True)
    with open(os.path.join(OUT, name + ".srt"), "w", encoding="utf-8-sig") as fh:
        fh.write("\n".join(srt))
    print("TOTAL", f"{t:.1f}s", out)

asyncio.run(main(*VIDEOS[sys.argv[1] if len(sys.argv) > 1 else "完整版"]))
