# Quick Start Guide - Task Manager v2.0

## ✅ Step 1: Run the SQL Script (REQUIRED!)

**Before running the application, you MUST create the database tables:**

1. Open **SQL Server Management Studio (SSMS)**
2. Connect to server: `10.195.103.198`
3. Select database: `TaskManagerDB`
4. Open file: `sql/create_enhanced_features.sql`
5. Click **Execute** (or press F5)
6. Verify success: You should see "Enhanced Features Schema Created Successfully!"

**Expected output:**
```
Created TimeEntries table
Created Subtasks table
Created TaskTemplates table
Created TemplateSubtasks table
Created TaskAnalytics table
Created SavedFilters table
Created composite search index on Tasks
Created date range index on Tasks
============================================
Enhanced Features Schema Created Successfully!
============================================
```

---

## ✅ Step 2: Run the Application

```cmd
python TaskManager_0.17.py
```

Or if you have the compiled version:
```cmd
TaskManager_0.17.exe
```

---

## ✅ Step 3: Test the Features

### 🔍 Test Time Tracking (Working!)

1. Click "📋 Task Manager" from main menu
2. **Right-click** any task
3. Select "▶ Start Timer"
4. Wait a minute
5. Right-click again → "⏹ Stop Timer"
6. You'll see how much time was logged!
7. Right-click → "⏱ View Time Log" to see history

**Expected behavior:**
- Timer starts and logs to database
- Timer stops and calculates duration
- Duration is saved in minutes
- Time log shows all entries for that task

---

### ✅ Test Subtasks (Working!)

1. Right-click any task
2. Select "✓ Manage Subtasks"
3. Click "Add Subtask"
4. Enter a name (e.g., "Step 1: Review design")
5. Double-click the subtask to mark it complete (checkbox changes to ☑)
6. Add more subtasks
7. See completion percentage update in header (e.g., "Completion: 50%")

**Expected behavior:**
- Subtasks dialog opens
- Can add/delete/toggle subtasks
- Completion percentage calculates automatically
- Main task list will show progress (coming soon)

---

### 📊 Test Dashboard (Placeholder)

1. From main menu, click "📊 Dashboard"
2. You'll see a "coming soon" message
3. **The database is ready!** The views are created and waiting
4. Just needs UI charts (can be added later)

**Database views created:**
- `vw_TaskSummary` - Overall metrics
- `vw_TaskSummaryByUser` - Per-user stats
- `vw_TimeTrackingSummary` - Time tracking data
- `vw_SubtaskProgress` - Subtask completion

**Test the views in SSMS:**
```sql
SELECT * FROM vw_TaskSummary;
SELECT * FROM vw_TaskSummaryByUser;
SELECT * FROM vw_TimeTrackingSummary;
```

---

### 📋 Test Templates (Placeholder)

1. From main menu, click "📋 Templates"
2. You'll see a "coming soon" message
3. **The database is ready!** Tables created
4. Just needs UI for creating/applying templates

**Database tables created:**
- `TaskTemplates` - Template configurations
- `TemplateSubtasks` - Pre-configured subtask lists

---

### 🔍 Test Advanced Search (Enhanced)

1. In task manager, type keywords in the search box
2. It searches across ALL fields now (title, description, vessel, drawing, etc.)
3. Click the 🔍 button next to search (shows info dialog)
4. **Full search UI coming soon**, but basic search is enhanced!

---

## 🎯 What's Working Right Now

### ✅ Fully Functional Features:
1. **Time Tracking** ⏱️
   - Start/stop timers
   - View time history
   - See total time logged
   - Manual time entries (via code)

2. **Subtasks** ✅
   - Add/delete/toggle subtasks
   - Calculate completion %
   - View subtask list
   - Double-click to complete

3. **Enhanced Search** 🔍
   - Full-text search (all fields)
   - Case-insensitive
   - Instant results

### 🚧 Database Ready, UI Coming Soon:
4. **Dashboard** 📊
   - All database views created ✅
   - Queries optimized ✅
   - Charts UI needed 🚧

5. **Templates** 📋
   - All database tables created ✅
   - Template manager class ready ✅
   - Create/apply UI needed 🚧

---

## 🔧 Troubleshooting

### Issue: "Feature Not Available" message
**Solution:** Run the SQL script (`sql/create_enhanced_features.sql`)

### Issue: Timer doesn't start
**Check:**
1. Database connection is working
2. TimeEntries table exists: `SELECT * FROM TimeEntries`
3. Check `task_manager.log` for errors

### Issue: Subtasks don't save
**Check:**
1. Subtasks table exists: `SELECT * FROM Subtasks`
2. Task ID is valid
3. Check `task_manager.log` for errors

### Issue: Application won't start
**Check:**
1. `task_manager_enhancements.py` is in the same folder
2. Check `task_manager.log` for import errors
3. Make sure Python can find the module

---

## 📝 Quick Reference

### Context Menu (Right-Click on Task):
- **View Details** - See full task info
- **Edit Task** - Modify task
- **▶ Start Timer** - Begin time tracking
- **⏹ Stop Timer** - End time tracking (if running)
- **⏱ View Time Log** - See time history
- **✓ Manage Subtasks** - Open subtask manager
- **Mark as Complete** - Complete task
- **Delete Task** - Remove task

### Main Menu Buttons:
- **📋 Task Manager** - Main task view
- **👥 User Management** - Manage users
- **⇄ Export/Import** - Excel import/export
- **📊 Dashboard** - Analytics (UI coming soon)
- **📋 Templates** - Task templates (UI coming soon)

### Search Box:
- Type keywords to search
- Searches: title, description, vessel, drawing, request, link
- Click **🔍** for advanced options (coming soon)
- Press **Escape** to clear search

---

## 🎓 Video Tutorial Ideas

If you want to train users, record these demos:

1. **"Time Tracking in 60 Seconds"**
   - Show start timer → work → stop timer → view log

2. **"Breaking Down Tasks with Subtasks"**
   - Open task → add 5 subtasks → check them off → show completion %

3. **"Finding Tasks Fast"**
   - Type in search box → instant results

---

## 📞 Support

**If something doesn't work:**
1. Check `task_manager.log` in the application folder
2. Verify SQL script was run successfully
3. Test in SSMS: `SELECT * FROM TimeEntries`
4. Contact IT support with log file

---

## 🚀 Next Steps

**For Full Feature Completion:**
1. Add Dashboard charts UI (matplotlib/plotly integration)
2. Add Template creation dialog
3. Add Advanced Search dialog with filters
4. Add keyboard shortcuts (Ctrl+T for timer, etc.)
5. Add status bar timer display
6. Add subtask progress % in main task list

**All backend code is ready!** Just needs UI wiring.

---

## 🎉 Success!

You now have:
- ✅ Time tracking working
- ✅ Subtasks working
- ✅ Enhanced search working
- ✅ Database fully set up
- ✅ 2 more features ready for UI completion

**Enjoy your upgraded Task Manager!**

---

**Version:** 2.0  
**Date:** January 22, 2026  
**Status:** 3/5 Features Fully Working, 2/5 Database Ready
