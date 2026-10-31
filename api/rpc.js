// Vercel Serverless Function: secure same-origin proxy to Google Apps Script.
// Set GAS_WEB_APP_URL in Vercel Project Settings > Environment Variables.
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, code: 'METHOD', message: 'Method not allowed' });
  }

  const gasUrl = process.env.GAS_WEB_APP_URL;
  if (!gasUrl || !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(gasUrl)) {
    return res.status(500).json({
      ok: false,
      code: 'CONFIG',
      message: 'ยังไม่ได้ตั้งค่า GAS_WEB_APP_URL ใน Vercel'
    });
  }

  try {
    const upstream = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        token: req.body && req.body.token || null,
        fn: String(req.body && req.body.fn || ''),
        args: Array.isArray(req.body && req.body.args) ? req.body.args : []
      }),
      redirect: 'follow',
      signal: AbortSignal.timeout(60000)
    });

    const raw = await upstream.text();
    let data;
    try {
      data = JSON.parse(raw);
    } catch (_) {
      return res.status(502).json({
        ok: false,
        code: 'UPSTREAM_FORMAT',
        message: 'Google Apps Script ตอบกลับไม่ใช่ JSON กรุณาตรวจสอบการ Deploy Web App'
      });
    }

    // Keep application errors in the normal RPC envelope; frontend reads {ok,data}/{ok:false}.
    return res.status(200).json(data);
  } catch (error) {
    return res.status(502).json({
      ok: false,
      code: 'UPSTREAM',
      message: 'เชื่อมต่อ Google Apps Script ไม่สำเร็จ: ' + (error.message || 'Unknown error')
    });
  }
};
