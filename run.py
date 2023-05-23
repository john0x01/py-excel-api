from api import create_app

app = create_app()
port = 2000

if __name__ == '__main__':
    app.run(port=port, debug=True)
    print("Server listening on port " + str(port))