# canon_bridge.py - Local Windows WIA Scanner Bridge for Canon MG2400
# Install requirements: pip install flask flask-cors pywin32 pillow
import base64
import io
import os
import tempfile
from flask import Flask, jsonify, request
from flask_cors import CORS
import pythoncom
import win32com.client
from PIL import Image

app = Flask(__name__)
CORS(app)

WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"

def find_wia_scanner():
    pythoncom.CoInitialize()
    manager = win32com.client.Dispatch("WIA.DeviceManager")
    for info in manager.DeviceInfos:
        if info.Type == 1:  # ScannerDeviceType
            name = info.Properties("Name").Value
            return info, name
    return None, None

@app.route("/status", methods=["GET"])
@app.route("/api/status", methods=["GET"])
def status():
    try:
        info, name = find_wia_scanner()
        if info:
            return jsonify({"status": "ok", "scanner_connected": True, "scanner_name": name})
        return jsonify({"status": "ok", "scanner_connected": False, "scanner_name": None})
    except Exception as e:
        return jsonify({"status": "error", "scanner_connected": False, "error": str(e)}), 500

@app.route("/scan", methods=["POST"])
@app.route("/api/scan", methods=["POST"])
def scan():
    payload = request.get_json(silent=True) or {}
    dpi = int(payload.get("dpi", 300))
    color_mode = payload.get("color_mode", "color")  # 'color' | 'grayscale' | 'bw'

    # WIA Current Intent: 1 = Color, 2 = Grayscale, 4 = Black & White Text
    intent_map = {"color": 1, "grayscale": 2, "bw": 4}
    wia_intent = intent_map.get(color_mode, 1)

    try:
        info, name = find_wia_scanner()
        if not info:
            return jsonify({"error": "Canon MG2400 scanner not detected via USB."}), 404

        device = info.Connect()
        item = device.Items[1]

        # Set WIA scan properties (6146 = Intent, 6147 = Horizontal DPI, 6148 = Vertical DPI)
        for prop_id, val in [(6146, wia_intent), (6147, dpi), (6148, dpi)]:
            try:
                item.Properties(str(prop_id)).Value = val
            except Exception:
                pass

        # Transfer image from flatbed via WIA PNG (with BMP fallback)
        tmp_path = os.path.join(tempfile.gettempdir(), "canon_mg2400_scan.tmp")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

        try:
            wia_image = item.Transfer(WIA_FORMAT_PNG)
        except Exception:
            wia_image = item.Transfer(WIA_FORMAT_BMP)

        wia_image.SaveFile(tmp_path)

        with Image.open(tmp_path) as img:
            rgb_img = img.convert("RGB")
            buf = io.BytesIO()
            rgb_img.save(buf, format="PNG")
            b64_str = base64.b64encode(buf.getvalue()).decode("utf-8")

        return jsonify({
            "image": f"data:image/png;base64,{b64_str}",
            "scanner_name": name
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("Canon MG2400 WIA Bridge running on http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)
