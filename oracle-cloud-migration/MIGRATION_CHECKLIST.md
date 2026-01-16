# HR Payroll System - Migration Checklist
## Google Sheets → Oracle Cloud MySQL

Use this checklist to track your migration progress.

---

## 📋 Pre-Migration (Preparation)

### Week Before Migration

- [ ] **Review current system**
  - [ ] Document all current workflows
  - [ ] List all users and their access levels
  - [ ] Identify peak usage times
  - [ ] Note any customizations or integrations

- [ ] **Backup everything**
  - [ ] Export Google Sheet to Excel (File → Download → Excel)
  - [ ] Save copy of all Apps Script code
  - [ ] Document all formulas and validations
  - [ ] Export all PDFs from Google Drive

- [ ] **Oracle Cloud account setup**
  - [ ] Create Oracle Cloud account
  - [ ] Verify credit card (won't be charged)
  - [ ] Choose home region wisely (can't change later)
  - [ ] Familiarize yourself with Oracle Cloud console

- [ ] **Prepare credentials**
  - [ ] Create service account for Google Sheets API
  - [ ] Download service account JSON key
  - [ ] Get Google Sheets spreadsheet ID
  - [ ] Choose strong MySQL password (20+ chars)

### Day Before Migration

- [ ] **Inform users**
  - [ ] Send notification about migration schedule
  - [ ] Request users to complete urgent tasks
  - [ ] Set expectations (system will be down 2-4 hours)

- [ ] **Final backup**
  - [ ] Take final snapshot of Google Sheet
  - [ ] Export all data to CSV
  - [ ] Record total counts (employees, payslips, loans)

---

## 🚀 Migration Day

### Phase 1: Oracle Cloud Setup (1-2 hours)

- [ ] **Create VM instance**
  - [ ] Log into Oracle Cloud console
  - [ ] Navigate to Compute → Instances
  - [ ] Click "Create Instance"
  - [ ] Name: `hr-payroll-server`
  - [ ] Shape: VM.Standard.A1.Flex (ARM)
  - [ ] OCPUs: 2, Memory: 12 GB
  - [ ] Image: Ubuntu 22.04 LTS
  - [ ] Generate SSH key pair
  - [ ] **Save private key securely!**
  - [ ] **Copy public IP address**

- [ ] **Configure firewall**
  - [ ] Add ingress rule for port 22 (SSH)
  - [ ] Add ingress rule for port 80 (HTTP)
  - [ ] Add ingress rule for port 443 (HTTPS)
  - [ ] Add ingress rule for port 3000 (API)

- [ ] **Connect to VM**
  - [ ] SSH into VM: `ssh -i key.pem ubuntu@YOUR_IP`
  - [ ] Confirm connection successful
  - [ ] Run `sudo apt update`

### Phase 2: Software Installation (30 minutes)

- [ ] **Install MySQL**
  ```bash
  sudo apt update
  sudo apt upgrade -y
  sudo apt install mysql-server -y
  sudo mysql_secure_installation
  ```
  - [ ] Set root password
  - [ ] Remove test database
  - [ ] Disallow root remote login

- [ ] **Install Node.js**
  ```bash
  curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
  sudo apt install -y nodejs
  node --version  # Verify
  ```

- [ ] **Install PM2**
  ```bash
  sudo npm install -g pm2
  ```

- [ ] **Configure Ubuntu firewall**
  ```bash
  sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 3000 -j ACCEPT
  sudo apt install iptables-persistent
  sudo netfilter-persistent save
  ```

### Phase 3: Database Setup (15 minutes)

- [ ] **Create database and user**
  ```bash
  sudo mysql
  ```
  ```sql
  CREATE DATABASE hrdb CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
  CREATE USER 'hrapp'@'localhost' IDENTIFIED BY 'YOUR_STRONG_PASSWORD';
  GRANT ALL PRIVILEGES ON hrdb.* TO 'hrapp'@'localhost';
  FLUSH PRIVILEGES;
  EXIT;
  ```

- [ ] **Upload and import schema**
  ```bash
  # From local machine:
  scp -i key.pem database/schema.sql ubuntu@YOUR_IP:~/

  # On server:
  mysql -u hrapp -p hrdb < ~/schema.sql
  ```

- [ ] **Verify schema**
  ```bash
  mysql -u hrapp -p hrdb
  ```
  ```sql
  SHOW TABLES;
  -- Should show: employees, mastersalary, employee_loans, etc.
  DESCRIBE employees;
  EXIT;
  ```

### Phase 4: Data Migration (30-60 minutes)

- [ ] **Prepare migration script**
  - [ ] Upload google credentials JSON to server
  - [ ] Create .env file with configuration
  - [ ] Install dependencies: `npm install`

- [ ] **Run migration**
  ```bash
  cd /path/to/scripts
  export SPREADSHEET_ID="your_id"
  export SERVICE_ACCOUNT_FILE="./google-credentials.json"
  export MYSQL_HOST="localhost"
  export MYSQL_USER="hrapp"
  export MYSQL_PASSWORD="your_password"
  node migrate-data.js
  ```

- [ ] **Verify migration**
  ```sql
  -- Count records
  SELECT COUNT(*) FROM employees;
  SELECT COUNT(*) FROM mastersalary;
  SELECT COUNT(*) FROM employee_loans;

  -- Compare with Google Sheets counts
  -- Spot check a few records for accuracy
  SELECT * FROM mastersalary WHERE record_number = 1;
  ```

- [ ] **Record migration results**
  - [ ] Employees migrated: _______
  - [ ] Payslips migrated: _______
  - [ ] Loans migrated: _______
  - [ ] Any errors: _______

### Phase 5: API Deployment (15 minutes)

- [ ] **Upload API files**
  ```bash
  # From local machine:
  scp -i key.pem -r api/* ubuntu@YOUR_IP:/var/www/hr-payroll/
  ```

- [ ] **Install dependencies**
  ```bash
  cd /var/www/hr-payroll
  npm install --production
  ```

- [ ] **Configure environment**
  ```bash
  nano .env
  ```
  - [ ] Set DB_HOST=localhost
  - [ ] Set DB_USER=hrapp
  - [ ] Set DB_PASSWORD
  - [ ] Set JWT_SECRET
  - [ ] Save file

- [ ] **Start API**
  ```bash
  pm2 start server.js --name hr-payroll-api
  pm2 save
  pm2 startup
  # Run the command it outputs
  ```

- [ ] **Test API**
  ```bash
  curl http://localhost:3000/health
  # Should return: {"status":"healthy","database":"connected"}

  curl http://localhost:3000/api/employees
  # Should return employee list
  ```

### Phase 6: Frontend Update (1-2 hours)

- [ ] **Update API base URL**
  - [ ] Find all HTML files
  - [ ] Add API configuration to each:
    ```javascript
    const API_BASE_URL = 'http://YOUR_ORACLE_IP:3000';
    ```

- [ ] **Update Dashboard.html**
  - [ ] Replace `google.script.run` calls
  - [ ] Test loading employees
  - [ ] Test loading recent payslips

- [ ] **Update PayrollForm.html**
  - [ ] Update create payslip function
  - [ ] Update edit payslip function
  - [ ] **CRITICAL:** Test edit - should be <1 second!

- [ ] **Update Reports.html**
  - [ ] Update list payslips function
  - [ ] Update filters
  - [ ] Test pagination

- [ ] **Update Loans.html**
  - [ ] Update get loan balance
  - [ ] Update loan history

- [ ] **Update EmployeeForm.html**
  - [ ] Update employee CRUD operations

---

## ✅ Testing Phase (1-2 hours)

### Critical Tests

- [ ] **Payslip Operations (MOST IMPORTANT)**
  - [ ] Create new payslip
  - [ ] **Edit existing payslip (Time this! Should be <1 sec)**
  - [ ] View payslip details
  - [ ] Delete payslip (if implemented)
  - [ ] Verify calculations match old system

- [ ] **Employee Operations**
  - [ ] List all employees
  - [ ] View employee details
  - [ ] Create new employee
  - [ ] Edit employee
  - [ ] Verify hourly rates

- [ ] **Loan Operations**
  - [ ] View current loan balance
  - [ ] View loan history
  - [ ] Create loan transaction
  - [ ] Verify balance calculations

- [ ] **Reports**
  - [ ] Generate salary report
  - [ ] Filter by employee
  - [ ] Filter by date range
  - [ ] Filter by employer
  - [ ] Test pagination

### Performance Tests

- [ ] **Speed tests** (compare with old system)
  - Old: Edit payslip = 2 minutes
  - New: Edit payslip = _____ seconds (should be <1!)
  - Old: Load dashboard = 30 seconds
  - New: Load dashboard = _____ seconds
  - Old: Get loan balance = 15 seconds
  - New: Get loan balance = _____ seconds

### Data Accuracy Tests

- [ ] **Spot check payslips**
  - [ ] Pick 5 random payslips
  - [ ] Compare all fields with Google Sheets
  - [ ] Verify calculations (gross, net, paid to account)
  - [ ] Check loan integration

- [ ] **Spot check employees**
  - [ ] Pick 5 random employees
  - [ ] Verify all details match
  - [ ] Check hourly rates

---

## 🔄 Parallel Running (1 Week)

- [ ] **Week 1: Both systems active**
  - [ ] Make Google Sheet read-only
  - [ ] All new entries go to MySQL system
  - [ ] Users report any issues
  - [ ] Monitor performance
  - [ ] Daily comparison of data

- [ ] **User feedback**
  - [ ] Survey users on new system
  - [ ] Document any issues
  - [ ] Make adjustments as needed

---

## 🎉 Go Live (After 1 Week)

- [ ] **Final verification**
  - [ ] All users trained
  - [ ] No critical bugs
  - [ ] Performance acceptable
  - [ ] Backups working

- [ ] **Switch over**
  - [ ] Archive Google Sheet (don't delete!)
  - [ ] Update bookmarks/links
  - [ ] Send announcement to all users
  - [ ] Disable Apps Script (keep read-only access)

- [ ] **Monitor**
  - [ ] Watch logs: `pm2 logs hr-payroll-api`
  - [ ] Monitor resource usage: `pm2 monit`
  - [ ] Check database size: `du -sh /var/lib/mysql/hrdb`

---

## 🛡️ Post-Migration

### First Day
- [ ] Monitor logs continuously
- [ ] Be available for user support
- [ ] Quick response to any issues

### First Week
- [ ] Daily health checks
- [ ] Verify automated backups running
- [ ] Monitor performance metrics
- [ ] Collect user feedback

### First Month
- [ ] Weekly database backups
- [ ] Review system logs
- [ ] Optimize slow queries (if any)
- [ ] Plan SSL certificate installation

### Ongoing
- [ ] Monthly backup tests (restore to test DB)
- [ ] Quarterly security updates
- [ ] Annual review of system performance

---

## 📊 Success Metrics

Record these metrics to measure success:

### Performance
- [ ] Payslip edit time: _______ (target: <1 second)
- [ ] Dashboard load time: _______ (target: <2 seconds)
- [ ] Report generation: _______ (target: <5 seconds)
- [ ] Loan balance query: _______ (target: <0.1 seconds)

### Reliability
- [ ] Uptime %: _______ (target: 99%+)
- [ ] Number of crashes: _______ (target: 0)
- [ ] Data integrity checks: _______ (target: 100% match)

### Cost
- [ ] Monthly hosting cost: _______ (target: $0)
- [ ] Time saved per week: _______ hours

### User Satisfaction
- [ ] User satisfaction score: _______/10
- [ ] Number of support tickets: _______
- [ ] Feature requests: _______

---

## 🆘 Rollback Plan (In Case of Emergency)

If critical issues occur:

### Immediate Steps
1. [ ] Notify all users
2. [ ] Make Google Sheet writable again
3. [ ] Stop PM2 process: `pm2 stop hr-payroll-api`
4. [ ] Document the issue

### Investigation
1. [ ] Check logs: `pm2 logs hr-payroll-api`
2. [ ] Check database: `mysql -u hrapp -p hrdb`
3. [ ] Identify root cause
4. [ ] Fix issue

### Retry
1. [ ] Test fix in isolated environment
2. [ ] Schedule new migration date
3. [ ] Communicate with users

---

## 📝 Notes & Issues

Use this space to record any issues encountered:

**Issue 1:**
- Date: _______
- Description: _______
- Resolution: _______

**Issue 2:**
- Date: _______
- Description: _______
- Resolution: _______

---

## ✅ Final Checklist

Before marking migration complete:

- [ ] All data migrated successfully
- [ ] All users trained on new system
- [ ] Performance targets met (<1 sec edits!)
- [ ] Automated backups configured and tested
- [ ] SSL certificate installed (optional)
- [ ] Documentation updated
- [ ] Google Sheet archived (not deleted)
- [ ] User satisfaction acceptable
- [ ] No critical bugs
- [ ] System stable for 1+ week

---

**Migration completed on: _________________**

**Signed off by: _________________**

**Notes: _________________________________**

---

## 🎊 Congratulations!

You've successfully migrated from Google Sheets to a professional, scalable, FREE MySQL-based system!

**Benefits achieved:**
- ✅ 120x faster operations
- ✅ No rate limiting
- ✅ $0/month cost
- ✅ Scalable to 100+ employees
- ✅ Professional infrastructure

**Next steps:**
- Consider adding mobile app
- Implement advanced reporting
- Add more automation
- Explore additional Oracle Cloud services

---

**Need help? Review the troubleshooting guide in `docs/ORACLE_CLOUD_SETUP.md`**
