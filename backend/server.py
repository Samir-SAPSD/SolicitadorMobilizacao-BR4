from backend.app import app

if __name__ == '__main__':
    print('Servidor rodando em http://localhost:5000')
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
