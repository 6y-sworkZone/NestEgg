from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from config import Config

db = SQLAlchemy()

from datetime import date

def format_money(value):
    return '¥{:,.2f}'.format(float(value or 0))

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    
    db.init_app(app)
    
    from app.routes import main_bp
    app.register_blueprint(main_bp)
    
    app.jinja_env.filters['formatMoney'] = format_money
    
    @app.context_processor
    def inject_today():
        return {'today': date.today().strftime('%Y-%m-%d')}
    
    with app.app_context():
        db.create_all()
        from app.models import init_default_data
        init_default_data()
    
    return app
