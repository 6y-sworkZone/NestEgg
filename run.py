from app import create_app
import random

app = create_app()

if __name__ == '__main__':
    port = random.randint(10000, 60000)
    print(f"Starting server on port {port}")
    app.run(debug=True, port=port, host='0.0.0.0')
