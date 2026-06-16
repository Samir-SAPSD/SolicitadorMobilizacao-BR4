from flask import Flask

from backend.api.routes import api_bp
from backend.config import FRONTEND_STATIC_DIR, FRONTEND_TEMPLATES_DIR
from backend.services import REPORTS_FOLDER, UPLOAD_FOLDER

app = Flask(
    __name__,
    template_folder=str(FRONTEND_TEMPLATES_DIR),
    static_folder=str(FRONTEND_STATIC_DIR),
)
app.json.ensure_ascii = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['REPORTS_FOLDER'] = REPORTS_FOLDER
app.register_blueprint(api_bp)


if __name__ == '__main__':
    print('Servidor rodando em http://localhost:5000')
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
