from datetime import datetime, date
from app import db

class Account(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    type = db.Column(db.String(50), nullable=False)
    balance = db.Column(db.Float, default=0.0)
    initial_balance = db.Column(db.Float, default=0.0)
    color = db.Column(db.String(20), default='#3498db')
    is_archived = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    bill_day = db.Column(db.Integer)
    payment_day = db.Column(db.Integer)
    credit_limit = db.Column(db.Float, default=0.0)
    
    transactions = db.relationship('Transaction', backref='account', lazy=True, foreign_keys='Transaction.account_id')
    reconciliations = db.relationship('AccountReconciliation', backref='account', lazy=True)

class AccountReconciliation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    actual_balance = db.Column(db.Float, nullable=False)
    system_balance = db.Column(db.Float, nullable=False)
    difference = db.Column(db.Float, nullable=False)
    notes = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class AccountBalanceHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    balance = db.Column(db.Float, nullable=False)
    record_date = db.Column(db.Date, default=date.today, nullable=False)

class Category(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    type = db.Column(db.String(20), nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey('category.id'))
    icon = db.Column(db.String(50))
    is_default = db.Column(db.Boolean, default=False)
    
    children = db.relationship('Category', backref=db.backref('parent', remote_side=[id]))
    transactions = db.relationship('Transaction', backref='category', lazy=True)

class Transaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, default=date.today, nullable=False)
    amount = db.Column(db.Float, nullable=False)
    type = db.Column(db.String(20), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('category.id'))
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    to_account_id = db.Column(db.Integer, db.ForeignKey('account.id'))
    note = db.Column(db.String(200))
    tags = db.Column(db.String(200))
    is_transfer = db.Column(db.Boolean, default=False)
    transfer_pair_id = db.Column(db.Integer)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    to_account = db.relationship('Account', foreign_keys=[to_account_id])

class RecurringTransaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    type = db.Column(db.String(20), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('category.id'))
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    frequency = db.Column(db.String(20), default='monthly')
    day_of_month = db.Column(db.Integer, default=1)
    start_date = db.Column(db.Date, default=date.today)
    next_run_date = db.Column(db.Date, default=date.today)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Budget(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    month = db.Column(db.String(7), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    category = db.relationship('Category', backref='budgets')

class BudgetTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('BudgetTemplateItem', backref='template', lazy=True)

class BudgetTemplateItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey('budget_template.id'), nullable=False)
    category_id = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)

class Investment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    code = db.Column(db.String(50))
    type = db.Column(db.String(50))
    buy_price = db.Column(db.Float, nullable=False)
    quantity = db.Column(db.Float, nullable=False)
    current_price = db.Column(db.Float)
    buy_date = db.Column(db.Date, default=date.today)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    account = db.relationship('Account', backref='investments')

class Dividend(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    investment_id = db.Column(db.Integer, db.ForeignKey('investment.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    date = db.Column(db.Date, default=date.today)
    note = db.Column(db.String(200))

class InvestmentNote(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    investment_id = db.Column(db.Integer, db.ForeignKey('investment.id'), nullable=False)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class SavingsGoal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    target_amount = db.Column(db.Float, nullable=False)
    current_amount = db.Column(db.Float, default=0.0)
    deadline = db.Column(db.Date)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'))
    priority = db.Column(db.Integer, default=1)
    is_completed = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    deposits = db.relationship('GoalDeposit', backref='goal', lazy=True)

class GoalDeposit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    goal_id = db.Column(db.Integer, db.ForeignKey('savings_goal.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    date = db.Column(db.Date, default=date.today)
    note = db.Column(db.String(200))

class ReportShare(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    share_code = db.Column(db.String(32), unique=True, nullable=False)
    report_type = db.Column(db.String(50), nullable=False)
    start_date = db.Column(db.Date)
    end_date = db.Column(db.Date)
    include_transactions = db.Column(db.Boolean, default=True)
    include_summary = db.Column(db.Boolean, default=True)
    include_budget = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime)

def init_default_data():
    if Category.query.count() == 0:
        income_categories = [
            ('工资', 'income', '💼'),
            ('奖金', 'income', '🎁'),
            ('投资收益', 'income', '📈'),
            ('兼职', 'income', '💻'),
            ('红包', 'income', '🧧'),
            ('其他收入', 'income', '💰')
        ]
        
        expense_categories = [
            ('餐饮', 'expense', '🍔'),
            ('交通', 'expense', '🚗'),
            ('购物', 'expense', '🛒'),
            ('娱乐', 'expense', '🎮'),
            ('居住', 'expense', '🏠'),
            ('医疗', 'expense', '💊'),
            ('教育', 'expense', '📚'),
            ('通讯', 'expense', '📱'),
            ('其他支出', 'expense', '📦')
        ]
        
        for name, type_, icon in income_categories + expense_categories:
            category = Category(name=name, type=type_, icon=icon, is_default=True)
            db.session.add(category)
        
        db.session.commit()
