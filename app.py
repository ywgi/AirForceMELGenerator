from flask import Flask
import os
from config import Config

# Import blueprints
from routes.upload_routes import upload_bp
from routes.main_routes import main_bp


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Ensure upload directory exists
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    # Register blueprints
    app.register_blueprint(main_bp)
    app.register_blueprint(upload_bp)

    return app


# Create app instance
app = create_app()

if __name__ == '__main__':
    app.run(debug=app.config['DEBUG'])