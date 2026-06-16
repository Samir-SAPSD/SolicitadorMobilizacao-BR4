from pathlib import Path
import sys

# Permite executar tanto com "python -m backend.server" quanto "python backend/server.py".
if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from backend.app import app

if __name__ == '__main__':
    print('Servidor rodando em http://localhost:5000')
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
