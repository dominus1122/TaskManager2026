# TaskManager Database Transfer Guide

## Overview
Your TaskManager application uses a SQL Server database called **TaskManagerDB** on server **10.195.103.198**.

---

## Tables Used by Your Application

### **Core Tables (Required)**

1. **Tasks** - Main task data
   - Contains all task information (title, description, dates, assignments, etc.)
   - Has soft delete column (`deleted = 0` for active tasks)

2. **Users** - User management
   - Stores user display names, TCNs, and SQL usernames
   - Has active/inactive status

3. **UserSessions** - Session tracking
   - Tracks who is currently using the application
   - Can be safely excluded if you don't need login history

4. **ScheduledTasks** - Work scheduler data
   - Stores task assignments per user per day
   - Links to Tasks table (foreign key: task_id)

5. **ScheduledHolidays** - Holiday/non-working days
   - Simple table with dates marked as holidays

### **Optional Tables (Enhanced Features)**
These only exist if you ran `sql/create_enhanced_features.sql`:

6. **TimeEntries** - Time tracking
7. **Subtasks** - Task checklists/subtasks
8. **TaskTemplates** - Reusable task templates
9. **TemplateSubtasks** - Template checklists
10. **TaskAnalytics** - Dashboard metrics cache
11. **SavedFilters** - Saved search filters

---

## Transfer Options

### Option 1: Copy Entire Tables
The simplest approach is to copy these tables to your new server:

**Minimum (Core functionality):**
- Tasks
- Users

**Standard (Recommended):**
- Tasks
- Users
- ScheduledTasks
- ScheduledHolidays

**Complete (Everything):**
- All tables listed above

### Option 2: Custom Query-Based Transfer
Use the queries in `sql/data_transfer_guide.sql` to:
- Export only active data (exclude deleted/inactive records)
- Export specific date ranges
- Export specific users or vessels
- Cherry-pick data based on your needs

---

## Transfer Order (Important!)

Due to foreign key relationships, transfer in this order:

1. **Users** (no dependencies)
2. **Tasks** (no dependencies)
3. **ScheduledTasks** (depends on Tasks)
4. **ScheduledHolidays** (no dependencies)
5. Optional tables if needed (most depend on Tasks)

---

## Quick Start

### For Entire Table Copy:
```sql
-- On source server (10.195.103.198):
SELECT * FROM TaskManagerDB.dbo.Tasks WHERE deleted = 0;
SELECT * FROM TaskManagerDB.dbo.Users WHERE active = 1;
SELECT * FROM TaskManagerDB.dbo.ScheduledTasks;
SELECT * FROM TaskManagerDB.dbo.ScheduledHolidays;
```

### For Specific Data Query:
```sql
-- Export tasks for a specific vessel
SELECT * FROM Tasks 
WHERE vessel_name = 'YOUR_VESSEL_NAME' 
  AND deleted = 0;

-- Export tasks created in last 6 months
SELECT * FROM Tasks 
WHERE created_date >= DATEADD(MONTH, -6, GETDATE())
  AND deleted = 0;

-- Export tasks assigned to specific users
SELECT * FROM Tasks 
WHERE assigned_to IN ('User1', 'User2', 'User3')
  AND deleted = 0;
```

---

## Files Created for You

I've created a comprehensive SQL file with all queries you need:

📄 **sql/data_transfer_guide.sql**
- Contains all SELECT queries for each table
- Shows table relationships
- Includes sample queries for custom data export
- Has row count queries to see how much data you have
- Includes schema information

---

## Recommendations

### For Testing:
1. Copy just **Tasks** and **Users** tables
2. Test your application on the new server
3. Verify all features work

### For Production:
1. Copy all core tables (Tasks, Users, ScheduledTasks, ScheduledHolidays)
2. Include enhanced feature tables if you use them
3. Keep UserSessions only if you need login history

### Things to Note:
- **deleted = 0**: Active tasks (set to 1 means deleted)
- **active = 1**: Active users (set to 0 means inactive)
- **IDENTITY columns**: All tables use auto-incrementing IDs
  - If copying with existing IDs: Use `SET IDENTITY_INSERT [table] ON`
  - If generating new IDs: Will need to update foreign keys

---

## Data Volume Check

Run this to see how much data you have:

```sql
-- See row counts
SELECT 'Tasks (Total)' AS Info, COUNT(*) AS Count FROM Tasks
UNION ALL
SELECT 'Tasks (Active)', COUNT(*) FROM Tasks WHERE deleted = 0
UNION ALL
SELECT 'Users (Total)', COUNT(*) FROM Users
UNION ALL
SELECT 'Users (Active)', COUNT(*) FROM Users WHERE active = 1
UNION ALL
SELECT 'ScheduledTasks', COUNT(*) FROM ScheduledTasks
UNION ALL
SELECT 'ScheduledHolidays', COUNT(*) FROM ScheduledHolidays;
```

---

## Need Help?

1. Open `sql/data_transfer_guide.sql` - it has all the queries you need
2. The queries are organized by:
   - Individual table exports
   - Complete data exports
   - Custom filtered exports
   - Relationship queries
3. Each query has comments explaining what it does

---

## Example: Full Standard Export

```sql
-- Step 1: Export Users
SELECT * FROM Users WHERE active = 1;

-- Step 2: Export Tasks (active only)
SELECT * FROM Tasks WHERE deleted = 0;

-- Step 3: Export Scheduling Data
SELECT * FROM ScheduledTasks;
SELECT * FROM ScheduledHolidays;
```

Save each result set to CSV or use SQL Server's import/export wizard.
