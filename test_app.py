from app import create_app

app = create_app()
client = app.test_client()

print("Testing API endpoints...")

# Test index page
response = client.get('/')
print(f"Index page: {response.status_code}")
assert response.status_code == 200
print("✓ Index page works")

# Test accounts page
response = client.get('/accounts')
print(f"Accounts page: {response.status_code}")
assert response.status_code == 200
print("✓ Accounts page works")

# Test transactions page
response = client.get('/transactions')
print(f"Transactions page: {response.status_code}")
assert response.status_code == 200
print("✓ Transactions page works")

# Test budgets page
response = client.get('/budgets')
print(f"Budgets page: {response.status_code}")
assert response.status_code == 200
print("✓ Budgets page works")

# Test investments page
response = client.get('/investments')
print(f"Investments page: {response.status_code}")
assert response.status_code == 200
print("✓ Investments page works")

# Test savings page
response = client.get('/savings')
print(f"Savings page: {response.status_code}")
assert response.status_code == 200
print("✓ Savings page works")

# Test analysis page
response = client.get('/analysis')
print(f"Analysis page: {response.status_code}")
assert response.status_code == 200
print("✓ Analysis page works")

# Test reports page
response = client.get('/reports')
print(f"Reports page: {response.status_code}")
assert response.status_code == 200
print("✓ Reports page works")

# Test API endpoints
response = client.get('/api/categories')
print(f"Categories API: {response.status_code}")
assert response.status_code == 200
print("✓ Categories API works")

response = client.get('/api/health-score')
print(f"Health score API: {response.status_code}")
assert response.status_code == 200
print("✓ Health score API works")

print("\n🎉 All tests passed!")
