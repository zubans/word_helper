import json
import urllib.request
import ssl


def _call_ollama(prompt: str) -> str:
    req = urllib.request.Request(
        "https://localhost/ollama/api/generate",
        data=json.dumps({
            "model": "gemma3:1b",
            "prompt": prompt,
            "stream": False
        }).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, context=ctx) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        return data.get("response", "")


def _get_active_document(ctx):
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.getCurrentComponent()


def Send(*args):
    ctx = None
    try:
        import uno
        ctx = uno.getComponentContext()
        doc = _get_active_document(ctx)
        if not doc:
            return None
        sel = doc.getCurrentController().getSelection()
        text = ""
        if sel and sel.getCount() > 0:
            text = sel.getByIndex(0).getString()
        if not text:
            text = doc.getText().getString()
        if not text:
            return None
        result = _call_ollama("Перепиши текст более ясно и кратко:\n" + text)
        if sel and sel.getCount() > 0:
            sel.getByIndex(0).setString(result)
        else:
            doc.getText().setString(result)
    except Exception:
        return None
    return None

