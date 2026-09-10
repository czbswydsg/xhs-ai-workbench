"""
小红书采集核心模块 (基于 Playwright)
- 自动打开浏览器，扫码登录（Cookie 持久化，下次免扫码）
- 根据关键词搜索，自动滑页采集笔记卡片
- 可选：进入详情页采集正文与收藏数
- 反爬：使用带界面的真实 Chromium，关闭 AutomationControlled 特征
"""
import json
import os
import random
import re
import time
import urllib.parse

from playwright.sync_api import sync_playwright

BASE_URL = "https://www.xiaohongshu.com"
COOKIE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cookies.json")
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")

# ---------- 数量解析：把 "1.2万" / "3.5k" / "1,234" 转成整数 ----------
def parse_count(text):
    if not text:
        return 0
    s = str(text).strip().replace(",", "").replace("+", "").replace(" ", "")
    if s in ("", "赞", "收藏", "评论", "点赞"):
        return 0
    try:
        if "万" in s:
            return int(float(s.replace("万", "")) * 10000)
        if "w" in s.lower():
            return int(float(s.lower().replace("w", "")) * 10000)
        if "千" in s or "k" in s.lower():
            return int(float(s.lower().replace("千", "").replace("k", "")) * 1000)
        return int(re.sub(r"[^\d]", "", s) or 0)
    except Exception:
        return 0


# ---------- 搜索卡片提取 JS ----------
EXTRACT_CARDS_JS = r"""
() => {
  const out = [];
  const cards = document.querySelectorAll('section.note-item');
  cards.forEach(card => {
    const coverA = card.querySelector('a.cover') || card.querySelector('a[href*="/search_result/"]');
    const href = coverA ? coverA.getAttribute('href') : '';
    const m = (href || '').match(/\/search_result\/([A-Za-z0-9]+)/);
    const id = m ? m[1] : '';
    if (!id) return;
    const titleEl = card.querySelector('.title') || card.querySelector('a.title');
    const authorEl = card.querySelector('.author .name') || card.querySelector('.name');
    const likeEl = card.querySelector('.like-wrapper .count') || card.querySelector('.count');
    // 封面图：优先 img src / data-src，其次背景图
    let cover = '';
    const coverImg = card.querySelector('a.cover img');
    if (coverImg) {
      cover = coverImg.getAttribute('src') || coverImg.getAttribute('data-src') || '';
    }
    if (!cover) {
      const style = coverA ? coverA.getAttribute('style') || '' : '';
      const cm = style.match(/url\(["']?(.*?)["']?\)/);
      if (cm) cover = cm[1];
    }
    out.push({
      id: id,
      title: titleEl ? titleEl.innerText.trim() : '',
      author: authorEl ? authorEl.innerText.trim() : '',
      likes_raw: likeEl ? likeEl.innerText.trim() : '',
      cover: cover,
      url: 'https://www.xiaohongshu.com' + href
    });
  });
  return out;
}
"""

# ---------- 详情页提取 JS（正文 + 收藏） ----------
EXTRACT_DETAIL_JS = r"""
() => {
  const out = {
    title: '', content: '', author: '', likes_raw: '', collects_raw: '',
    comments_raw: '', shares_raw: '', fans_raw: '', publish_time: '', images: []
  };

  // ===== 1) 优先从页面内嵌的结构化数据取（最稳，不受 class 改名影响）=====
  try {
    let state = window.__INITIAL_STATE__;
    if (typeof state === 'string') {
      try { state = JSON.parse(state); } catch (e) { state = null; }
    }
    let note = null;
    const tryMap = (m) => {
      if (!m) return null;
      for (const k of Object.keys(m)) {
        const v = m[k];
        if (v && v.note) return v.note;
      }
      return null;
    };
    if (state) {
      if (state.note && state.note.noteDetailMap) note = tryMap(state.note.noteDetailMap);
      else if (state.noteDetailMap) note = tryMap(state.noteDetailMap);
      else if (state.note && state.note.note) note = state.note.note;
    }
    if (note) {
      const ii = note.interactInfo || {};
      if (ii.liked_count != null) out.likes_raw = String(ii.liked_count);
      if (ii.collected_count != null) out.collects_raw = String(ii.collected_count);
      if (ii.comment_count != null) out.comments_raw = String(ii.comment_count);
      if (ii.share_count != null) out.shares_raw = String(ii.share_count);
      out.title = note.title || '';
      out.content = note.desc || note.content || '';
      if (note.user) {
        out.author = note.user.nickname || note.user.name || '';
        if (note.user.fans != null) out.fans_raw = String(note.user.fans);
        else if (note.user.fans_count != null) out.fans_raw = String(note.user.fans_count);
      }
      if (note.time) {
        const t = Number(note.time);
        if (!isNaN(t)) {
          const dt = new Date(t * 1000);
          if (!isNaN(dt.getTime())) out.publish_time = dt.toISOString().slice(0, 10);
        }
      }
      const imgList = note.imageList || [];
      imgList.forEach(im => {
        let u = im.url ||
               (im.infoList && im.infoList.length && im.infoList[im.infoList.length - 1].url) ||
               im.urlDefault || '';
        if (u && !u.startsWith('data:')) out.images.push(u.split('?')[0]);
      });
      if (!out.images.length && note.video && note.video.media && note.video.media.url) {
        out.images.push(note.video.media.url.split('?')[0]);
      }
    }
  } catch (e) { /* 忽略，走 DOM 兜底 */ }

  // ===== 2) DOM 兜底（仅在结构化数据缺失时）=====
  const pick = (sel) => { const e = document.querySelector(sel); return e ? e.innerText.trim() : ''; };
  if (!out.title) out.title = pick('#detail-title') || pick('.title') || pick('h1');
  if (!out.content) out.content = pick('#detail-desc') || pick('.desc') || pick('.content');
  if (!out.author) out.author = pick('.author-wrapper .name') || pick('.username');

  const bar = document.querySelector('.engage-bar') || document.querySelector('.interact-container');
  if (bar) {
    const bc = bar.querySelectorAll('.count');
    if (bc.length >= 4) {
      if (!out.likes_raw) out.likes_raw = bc[0].innerText.trim();
      if (!out.collects_raw) out.collects_raw = bc[1].innerText.trim();
      if (!out.comments_raw) out.comments_raw = bc[2].innerText.trim();
      if (!out.shares_raw) out.shares_raw = bc[3].innerText.trim();
    } else {
      if (!out.likes_raw) out.likes_raw = pick('.like-wrapper .count');
      if (!out.collects_raw) out.collects_raw = pick('.collect-wrapper .count');
      if (!out.comments_raw) out.comments_raw = pick('.comment-wrapper .count');
      if (!out.shares_raw) out.shares_raw = pick('.share-wrapper .count');
    }
  } else {
    if (!out.likes_raw) out.likes_raw = pick('.like-wrapper .count');
    if (!out.collects_raw) out.collects_raw = pick('.collect-wrapper .count');
    if (!out.comments_raw) out.comments_raw = pick('.comment-wrapper .count');
    if (!out.shares_raw) out.shares_raw = pick('.share-wrapper .count');
  }
  if (!out.fans_raw) {
    const f = document.querySelector('.author-wrapper .follower .count') ||
              document.querySelector('.author-wrapper .count');
    if (f) out.fans_raw = f.innerText.trim();
  }
  if (!out.publish_time) {
    const d = document.querySelector('#detail-date') ||
              document.querySelector('.publish-date') ||
              document.querySelector('time') ||
              document.querySelector('.date');
    if (d) out.publish_time = (d.getAttribute('datetime') || d.innerText.trim()).slice(0, 10);
  }
  if (!out.images.length) {
    const imgs = document.querySelectorAll(
      '#detail-media img, .note-slider img, .swiper-slide img, .detail-image img, .img-container img'
    );
    imgs.forEach(img => {
      let u = img.getAttribute('data-src') || img.getAttribute('src') || '';
      if (!u && img.getAttribute('srcset')) u = img.getAttribute('srcset').split(',')[0].trim().split(' ')[0];
      if (u && !u.startsWith('data:')) out.images.push(u.split('?')[0]);
    });
  }
  return out;
}
"""


class XHSScraper:
    def __init__(self, headless=False, login_timeout=180, progress=None, cookie_file=None,
                 qr_capture_path=None):
        """
        progress: 可选回调 progress(msg) 用于输出进度
        cookie_file: 登录态 cookie 的存取路径。默认用全局 cookies.json；
                     传入自定义路径时（如 sessions/{visitor}.json）实现"每人独立登录态"，
                     不传则与旧行为完全一致（多用户扩展点）。
        qr_capture_path: 可选；在无头模式等待扫码时，把登录页（含二维码）截图持续写到
                     该路径，供界面（如 Streamlit st.image）展示后由用户手机扫码。
        """
        self.headless = headless
        self.login_timeout = login_timeout
        self.progress = progress or (lambda m: None)
        self._cookie_file = cookie_file or COOKIE_FILE
        self.qr_capture_path = qr_capture_path
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        # 为 True 时 launch() 不加载旧 Cookie（强制扫码登录用）
        self._skip_cookies = False

    def _log(self, msg):
        self.progress(msg)

    def _get_web_session(self):
        """返回当前 context 中的 web_session cookie（小红书真实登录态凭证），没有则返回 None。"""
        try:
            for c in self.context.cookies():
                if c.get("name") == "web_session" and c.get("value"):
                    return c
        except Exception:
            pass
        return None

    def launch(self):
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(
            headless=self.headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )
        self.context = self.browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )
        # 加载已保存的 Cookie，免扫码（强制扫码登录时跳过，避免旧登录态干扰）
        if os.path.exists(self._cookie_file) and not self._skip_cookies:
            try:
                with open(self._cookie_file, "r", encoding="utf-8") as f:
                    cookies = json.load(f)
                self.context.add_cookies(cookies)
                self._log("已加载已保存的登录态，尝试免扫码登录")
            except Exception as e:
                self._log(f"Cookie 加载失败: {e}")
        self.page = self.context.new_page()
        self.page.set_default_timeout(30000)

    def _save_cookies(self):
        try:
            cookies = self.context.cookies()
            d = os.path.dirname(self._cookie_file)
            if d:
                os.makedirs(d, exist_ok=True)
            with open(self._cookie_file, "w", encoding="utf-8") as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            self._log("登录态已保存，下次可免扫码")
        except Exception as e:
            self._log(f"Cookie 保存失败: {e}")

    def _wait_qr_login(self, old_value="", poll_interval=2):
        """等待用户扫码，直到出现新的 web_session cookie。超时抛 TimeoutError。
        poll_interval: 每次轮询间隔秒数（默认 2s，给浏览器 UI 充足响应时间，避免卡顿）。

        云端无头修复：登录页用 domcontentloaded 快速加载（networkidle 在海外访问
        小红书时经常 60s 超时，导致二维码迟迟截不出来）；截图只在已进入登录页时
        进行，避免把首页/空白页当二维码展示；关键步骤都写入日志便于界面排查。
        """
        self._log("请用手机小红书 App 扫码登录（云端无头模式：二维码显示在页面下方）")
        opened_login = False
        login_try = 0
        deadline = time.time() + self.login_timeout
        while time.time() < deadline:
            cur = (self._get_web_session() or {}).get("value", "")
            if cur and cur != old_value:
                # 等其余 cookie 写入完成再保存
                self.page.wait_for_timeout(1500)
                self._log("扫码登录成功")
                return True
            # 页面上没有登录弹窗时，主动打开登录页方便扫码
            if not opened_login:
                has_modal = self.page.evaluate(
                    "() => { const el = document.querySelector('.login-container') "
                    "|| document.querySelector('#login-container'); "
                    "return !!el && el.offsetParent !== null; }"
                )
                if not has_modal:
                    try:
                        if login_try == 0:
                            self._log("正在打开小红书登录页，准备截取登录二维码…")
                        self.page.goto(BASE_URL + "/login",
                                       wait_until="domcontentloaded",
                                       timeout=30000)
                        self.page.wait_for_timeout(2500)
                        opened_login = True
                        self._log("已打开小红书登录页，二维码生成中…")
                    except Exception:
                        login_try += 1
                        if login_try >= 3:
                            self._log("打开登录页较慢，自动重试中…")
                            login_try = 0
            # 定期把浏览器置前，避免窗口被遮挡在后台看不到
            try:
                self.page.bring_to_front()
            except Exception:
                pass
            # 无头模式（云端）下用户看不到浏览器窗口：把登录页（含二维码）持续
            # 截图到 qr_capture_path，由界面 st.image 展示，用户手机扫码即可登录。
            if self.qr_capture_path:
                try:
                    url_now = self.page.url or ""
                    # 只截登录页，避免把首页/空白页当二维码展示
                    if opened_login or "login" in url_now:
                        self.page.screenshot(path=self.qr_capture_path)
                        if not getattr(self, "_qr_notified", False):
                            self._log("📱 二维码已就绪：请用手机小红书 App 扫码登录")
                            self._qr_notified = True
                except Exception:
                    pass
            self.page.wait_for_timeout(poll_interval * 1000)
        self._log("登录等待超时，可点界面「取消登录」或重新发起")
        raise TimeoutError("登录等待超时（未检测到新的 web_session）")

    def ensure_login(self):
        """打开首页；仅当存在 web_session（真实登录态）且页面无登录遮罩时才跳过扫码。

        说明：不能只靠 DOM 判断 —— 未登录时页面不一定弹登录框，
        会误判为"已登录"并保存无效 cookie，导致采集失败。
        """
        self.page.goto(BASE_URL, wait_until="domcontentloaded")
        self.page.wait_for_timeout(2000)
        old_value = (self._get_web_session() or {}).get("value", "")
        has_modal = self.page.evaluate(
            "() => { const el = document.querySelector('.login-container') "
            "|| document.querySelector('#login-container'); "
            "return !!el && el.offsetParent !== null; }"
        )
        if old_value and not has_modal and "login" not in (self.page.url or ""):
            self._log("已处于登录状态（检测到有效 web_session）")
            self._save_cookies()
            return
        self._log("检测到未登录")
        self._wait_qr_login(old_value=old_value)
        self._save_cookies()

    def login_only(self):
        """强制扫码登录并保存 Cookie（不采集），供界面「扫码登录」按钮调用。

        修复要点：
          1) 使用 networkidle 等待登录页完整加载；
          2) 显式等待 QR 码元素可见，再开始轮询 cookie（避免浏览器
             "一闪而过"——以前是加载没完成就已开始倒计时）;
          3) 轮询间隔 2s，留出充足 UI 响应时间，并定期 bring_to_front 防止窗口被挡；
          4) 超时后将浏览器再保留 60s 供用户补扫，期间检测到登录会立即成功；
          5) 超时到点才真正关浏览器，给用户充裕的扫码窗口。
        """
        self._skip_cookies = True
        self.launch()
        try:
            try:
                self.page.goto(BASE_URL + "/login", wait_until="networkidle",
                               timeout=60000)
            except Exception as e:
                self._log(f"登录页加载中（{e.__class__.__name__}），将稍后重试…")
                try:
                    self.page.goto(BASE_URL + "/login", wait_until="domcontentloaded",
                                   timeout=60000)
                except Exception:
                    pass
            self.page.wait_for_timeout(1500)
            # 显式等待 QR 元素渲染（多种选择器都试一遍）
            qr_selectors = (
                ".qrcode img, .qrcode-img, .login-qrcode img, "
                "#qrcode, .qr-code img, "
                "img[alt*='二维码'], img[src*='qrcode']"
            )
            try:
                self.page.wait_for_selector(qr_selectors, state="visible", timeout=20000)
                self._log("✅ 二维码已就绪，请在弹出的浏览器中用手机小红书 App 扫码")
            except Exception:
                # 拿不到具体元素也别急着失败 —— 部分版本二维码嵌在 iframe 里
                self.page.wait_for_timeout(3000)
                self._log("⚠️ 暂未识别到二维码元素，请在浏览器窗口中直接查看并扫码")

            # 把浏览器窗口带到前台（用户很多时候看不到后台窗口）
            try:
                self.page.bring_to_front()
            except Exception:
                pass

            # 关键：登录页刚加载完，小红书会立刻下发一个“匿名 web_session”。
            # 必须先把它的值记为基线 old_value——真正扫码成功后该值会轮换；
            # 若以 old_value="" 等待，会把匿名会话误判成“扫码登录成功”。
            _anon = (self._get_web_session() or {}).get("value", "")
            self._log("已记录当前会话基线，等待手机扫码后登录态轮换…")
            try:
                self._wait_qr_login(old_value=_anon, poll_interval=2)
            except TimeoutError:
                self._log("等待超时。浏览器窗口会再保留 60 秒，期间补扫也能识别登录…")
                # 把窗口再带回前台，给用户最后一次补扫机会
                try:
                    self.page.bring_to_front()
                except Exception:
                    pass
                grace_deadline = time.time() + 60
                while time.time() < grace_deadline:
                    cur = (self._get_web_session() or {}).get("value", "")
                    # 必须仍与匿名基线不同才算真正登录（匿名 cookie 全程存在，不能当作补扫成功）
                    if cur and cur != _anon:
                        self._log("补扫成功")
                        break
                    try:
                        self.page.bring_to_front()
                    except Exception:
                        pass
                    self.page.wait_for_timeout(2000)
                else:
                    raise
        finally:
            self.close()

    def login_with_qr_capture(self, qr_png_path):
        """访客扫码登录（二维码图片内嵌在界面展示，供远程访问者用自己的小红书账号登录）。

        流程：打开登录页 → 把二维码元素截图到 qr_png_path（供界面 st.image 展示）
        → 等待扫码成功（登录态写入 self._cookie_file，每人独立）→ 返回 True。
        超时抛 TimeoutError；浏览器窗口在扫码期间停留，结束自动关闭。
        """
        self._skip_cookies = True
        self.launch()
        try:
            try:
                self.page.goto(BASE_URL + "/login", wait_until="networkidle",
                               timeout=60000)
            except Exception:
                try:
                    self.page.goto(BASE_URL + "/login", wait_until="domcontentloaded",
                                   timeout=60000)
                except Exception:
                    pass
            self.page.wait_for_timeout(1500)
            qr_selectors = (
                ".qrcode img, .qrcode-img, .login-qrcode img, "
                "#qrcode, .qr-code img, "
                "img[alt*='二维码'], img[src*='qrcode']"
            )
            qr_el = None
            try:
                qr_el = self.page.wait_for_selector(qr_selectors, state="visible",
                                                    timeout=15000)
                self._log("✅ 二维码已就绪，请用手机小红书 App 扫码（二维码也已显示在工作台页面上）")
            except Exception:
                self._log("⚠️ 未识别到二维码元素，将截图登录区域供扫码")
            d = os.path.dirname(qr_png_path)
            if d:
                os.makedirs(d, exist_ok=True)
            if qr_el:
                try:
                    qr_el.screenshot(path=qr_png_path)
                    self._log("二维码图片已生成")
                except Exception:
                    self.page.screenshot(path=qr_png_path)
                    self._log("二维码元素截图失败，已截取整个登录页")
            else:
                self.page.screenshot(path=qr_png_path)
                self._log("已截取登录页（含二维码区域）")
            try:
                self.page.bring_to_front()
            except Exception:
                pass
            # 等待扫码成功：以“登录页自动下发的匿名 web_session”为基线，
            # 扫码成功才会轮换成新值 → 避免把匿名会话误判为已登录。
            _anon = (self._get_web_session() or {}).get("value", "")
            self._wait_qr_login(old_value=_anon, poll_interval=2)
            self._save_cookies()
            self._log("访客账号登录成功，登录态已保存到独立文件")
            return True
        except TimeoutError:
            self._log("扫码等待超时，请重新发起登录")
            raise
        finally:
            self.close()

    def search(self, keyword):
        url = BASE_URL + "/search_result?keyword=" + urllib.parse.quote(keyword)
        self.page.goto(url, wait_until="domcontentloaded")
        self.page.wait_for_timeout(2500)
        # 若搜索页弹出登录遮罩，等待其消失
        try:
            self.page.wait_for_function(
                "() => { const el = document.querySelector('.login-container') "
                "|| document.querySelector('#login-container'); "
                "return !el || el.offsetParent === null; }",
                timeout=self.login_timeout * 1000,
            )
        except Exception:
            pass

    def scroll_and_collect(self, target_count=50, max_scrolls=20):
        """自动滑页采集笔记卡片，去重"""
        collected = {}
        no_new_rounds = 0
        for i in range(max_scrolls):
            # 多种方式滚动，兼容不同页面结构
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            self.page.mouse.wheel(0, 1200)
            self.page.wait_for_timeout(1500 + random.randint(0, 800))
            cards = self.page.evaluate(EXTRACT_CARDS_JS)
            before = len(collected)
            for c in cards:
                collected[c["id"]] = c
            after = len(collected)
            self._log(f"已滑页 {i+1} 次，累计采集 {after} 条")
            if after >= target_count:
                self._log(f"已达到目标采集量 {target_count} 条")
                break
            if after == before:
                no_new_rounds += 1
                if no_new_rounds >= 3:
                    self._log("连续多次无新内容，停止滑页")
                    break
            else:
                no_new_rounds = 0
        return list(collected.values())

    def fetch_details(self, notes, limit=None, download_images=False):
        """逐个打开详情页，补充正文、收藏数与原图（较慢，注意反爬）"""
        if limit:
            notes = notes[:limit]
        detail_page = self.context.new_page()
        detail_page.set_default_timeout(30000)
        total = len(notes)
        for idx, note in enumerate(notes, 1):
            try:
                detail_page.goto(note["url"], wait_until="domcontentloaded")
                detail_page.wait_for_timeout(1200 + random.randint(0, 1000))
                data = detail_page.evaluate(EXTRACT_DETAIL_JS)
                if data.get("title"):
                    note["title"] = data["title"]
                note["content"] = data.get("content", "")
                note["collects_raw"] = data.get("collects_raw", "")
                note["comments_raw"] = data.get("comments_raw", "")
                note["shares_raw"] = data.get("shares_raw", "")
                note["fans_raw"] = data.get("fans_raw", "")
                note["publish_time"] = data.get("publish_time", "")
                note["images"] = data.get("images", [])
                note["likes"] = parse_count(data.get("likes_raw") or note.get("likes_raw"))
                note["collects"] = parse_count(data.get("collects_raw"))
                note["comments"] = parse_count(data.get("comments_raw"))
                note["shares"] = parse_count(data.get("shares_raw"))
                note["fans"] = parse_count(data.get("fans_raw"))
                # 高清无水印原图下载（走浏览器上下文，带登录态）
                local = []
                if download_images and note["images"]:
                    folder = os.path.join(DOWNLOAD_DIR, note["id"])
                    os.makedirs(folder, exist_ok=True)
                    for i, u in enumerate(note["images"], 1):
                        try:
                            resp = detail_page.goto(u, wait_until="load", timeout=20000)
                            body = resp.body() if resp else b""
                            if body:
                                with open(os.path.join(folder, f"{i}.jpg"), "wb") as fp:
                                    fp.write(body)
                                local.append(os.path.join(folder, f"{i}.jpg"))
                        except Exception as ex:
                            self._log(f"  图片下载失败: {ex}")
                        time.sleep(random.uniform(0.3, 0.8))
                    note["local_images"] = local
                    self._log(f"已下载 {len(local)} 张原图")
                self._log(f"详情 {idx}/{total}：{note.get('title','')[:20]} | 赞{note['likes']} 藏{note['collects']} | 图{len(note['images'])}")
            except Exception as e:
                note.setdefault("content", "")
                note.setdefault("images", [])
                note["collects"] = parse_count(note.get("collects_raw", ""))
                note["likes"] = parse_count(note.get("likes_raw", ""))
                self._log(f"详情 {idx}/{total} 失败: {e}")
            time.sleep(random.uniform(0.6, 1.6))
        detail_page.close()
        return notes

    def download_images_for(self, notes):
        """批量下载指定笔记的高清无水印原图（复用登录态，不重新采集）。

        notes 中已有 local_images 的图片会跳过；缺 images 字段的笔记会先打开详情页提取图片链接。
        返回更新后的 notes 与成功下载图片总数。
        """
        self.launch()
        total_imgs = 0
        try:
            self.ensure_login()
            page = self.context.new_page()
            page.set_default_timeout(30000)
            total = len(notes)
            for idx, note in enumerate(notes, 1):
                try:
                    imgs = note.get("images") or []
                    if not imgs and note.get("url"):
                        page.goto(note["url"], wait_until="domcontentloaded")
                        page.wait_for_timeout(800 + random.randint(0, 600))
                        data = page.evaluate(EXTRACT_DETAIL_JS)
                        imgs = data.get("images", [])
                        note["images"] = imgs
                    local = list(note.get("local_images") or [])
                    if imgs:
                        folder = os.path.join(DOWNLOAD_DIR, note["id"])
                        os.makedirs(folder, exist_ok=True)
                        for i, u in enumerate(imgs, 1):
                            p = os.path.join(folder, f"{i}.jpg")
                            if os.path.exists(p):
                                if p not in local:
                                    local.append(p)
                                continue
                            try:
                                resp = page.goto(u, wait_until="load", timeout=20000)
                                body = resp.body() if resp else b""
                                if body:
                                    with open(p, "wb") as fp:
                                        fp.write(body)
                                    local.append(p)
                                    total_imgs += 1
                            except Exception as ex:
                                self._log(f"  图片下载失败: {ex}")
                            time.sleep(random.uniform(0.3, 0.8))
                    note["local_images"] = list(dict.fromkeys(local))
                    self._log(f"图片 {idx}/{total}：{(note.get('title') or '')[:20]} | 本地 {len(note['local_images'])} 张")
                except Exception as e:
                    self._log(f"图片 {idx}/{total} 失败: {e}")
                time.sleep(random.uniform(0.4, 1.0))
            page.close()
            return notes, total_imgs
        finally:
            try:
                self.close()
            except Exception:
                pass

    def run(self, keyword, target_count=50, max_scrolls=20, with_details=True,
            detail_limit=None, download_images=False):
        self.launch()
        try:
            self.ensure_login()
            self.search(keyword)
            notes = self.scroll_and_collect(target_count=target_count, max_scrolls=max_scrolls)
            for n in notes:
                n.setdefault("likes", parse_count(n.get("likes_raw", "")))
                n.setdefault("collects", 0)
                n.setdefault("content", "")
                n.setdefault("collects_raw", "")
                n.setdefault("comments_raw", "")
                n.setdefault("comments", 0)
                n.setdefault("cover", "")
                n.setdefault("images", [])
                n.setdefault("local_images", [])
                n.setdefault("shares", 0)
                n.setdefault("fans", 0)
                n.setdefault("publish_time", "")
                n.setdefault("shares_raw", "")
                n.setdefault("fans_raw", "")
            if with_details:
                self._log("开始采集详情（正文+收藏+原图），请勿关闭浏览器…")
                notes = self.fetch_details(notes, limit=detail_limit, download_images=download_images)
            else:
                for n in notes:
                    n["collects"] = parse_count(n.get("collects_raw", ""))
            self._log(f"采集完成，共 {len(notes)} 条")
            return notes
        finally:
            try:
                self.close()
            except Exception:
                pass

    def close(self):
        try:
            if self.page:
                self.page.close()
        except Exception:
            pass
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.pw:
                self.pw.stop()
        except Exception:
            pass


if __name__ == "__main__":
    kw = input("输入搜索关键词: ")
    s = XHSScraper()
    res = s.run(kw, target_count=20, max_scrolls=10, with_details=True)
    print(json.dumps(res, ensure_ascii=False, indent=2))
