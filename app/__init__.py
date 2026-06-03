from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text
from config import Config

db = SQLAlchemy()

from datetime import date

def format_money(value):
    return '¥{:,.2f}'.format(float(value or 0))

def migrate_database(app):
    with app.app_context():
        inspector = inspect(db.engine)
        
        from app.models import ReportShare
        
        existing_columns = [c['name'] for c in inspector.get_columns('report_share')]
        expected_columns = [
            'include_transactions',
            'include_summary', 
            'include_budget'
        ]
        
        for col in expected_columns:
            if col not in existing_columns:
                print(f"Adding column: {col}")
                db.session.execute(text(f"ALTER TABLE report_share ADD COLUMN {col} BOOLEAN DEFAULT 1"))
        
        db.session.commit()

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    
    db.init_app(app)
    
    from app.routes import main_bp
    app.register_blueprint(main_bp)
    
    app.jinja_env.filters['formatMoney'] = format_money
    
    @app.context_processor
    def inject_today():
        from datetime import date as date_cls
        d = date_cls.today()
        return {'today': d.strftime('%Y-%m-%d'), 'current_month': d.strftime('%Y-%m')}
    
    with app.app_context():
        db.create_all()
        migrate_database(app)
        from app.models import init_default_data
        init_default_data()
    
    return app
