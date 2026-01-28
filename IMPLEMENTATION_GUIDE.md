# Task Manager Enhanced Features Implementation Guide

## Overview
This document provides step-by-step instructions for implementing the 5 major features for the Task Manager application.

## Database Setup
**CRITICAL: Run this first before using any new features!**

1. Open SQL Server Management Studio (SSMS)
2. Connect to your server: `10.195.103.198`
3. Open the script: `sql/create_enhanced_features.sql`
4. Execute the script against your `TaskManagerDB` database
5. Verify all tables were created:
   - TimeEntries
   - Subtasks
   - TaskTemplates
   - TemplateSubtasks
   - TaskAnalytics
   - SavedFilters

## Features Implemented

### 1. Task Time Tracking ⏱️
**Location:** Integrated into main TaskManager class

**How to use:**
- Click on a task → "Start Timer" button appears in context menu
- Timer runs in background, updates every minute
- Click "Stop Timer" to log the session
- View time history in "Task Details" dialog
- See total time logged in main task list

**Database:**
- Table: `TimeEntries`
- Tracks: start_time, end_time, duration_minutes, user_name
- Supports: Manual entry, auto-timer, and bulk edits

**Benefits:**
- Accurate billable hours tracking
- Compare estimated vs. actual time
- Identify time-consuming tasks
- Generate weekly timesheets

### 2. Subtasks/Checklists ✅
**Location:** Available in Task Edit dialog

**How to use:**
- Open any task → Click "Subtasks" button
- Add checklist items with descriptions
- Mark items complete as you progress
- See % completion in main task list (e.g., "65% - 13/20 items")
- Drag to reorder subtasks

**Database:**
- Table: `Subtasks`
- Tracks: title, description, completed, sort_order, assigned_to
- Cascades: Deletes with parent task

**Benefits:**
- Break complex tasks into steps
- Track partial progress
- Assign subtasks to team members
- Clear next actions

### 3. Dashboard/Analytics 📊
**Location:** Main Menu → "Dashboard" button

**How to use:**
- Click "Dashboard" from main menu
- View charts: Completion trends, workload by user, overdue tasks
- Filter by date range (This Week, This Month, This Quarter, Custom)
- Export reports to PDF/Excel
- Real-time metrics update every 5 minutes

**Database:**
- Views: `vw_TaskSummary`, `vw_TaskSummaryByUser`, `vw_TimeTrackingSummary`
- Cache: `TaskAnalytics` (for historical trends)

**Metrics shown:**
- Tasks by status (Active/Completed/Overdue)
- Completion rate %
- Average completion time
- Workload distribution
- Time logged vs. estimated
- Overdue task alert list

**Benefits:**
- Management visibility
- Identify bottlenecks
- Resource allocation insights
- Historical trends

### 4. Task Templates 📋
**Location:** Main Menu → "Templates" or Add Task → "Use Template"

**How to use:**
- Create template: Click "New Template" → Fill in default values
- Add template subtasks (pre-configured checklists)
- Use template: Click "Add Task" → "From Template" → Select → Auto-fills all fields
- Edit templates anytime
- Share templates with team (Public checkbox)

**Database:**
- Tables: `TaskTemplates`, `TemplateSubtasks`
- Tracks: template_name, default values, usage_count
- Supports: Public/Private templates per user

**Examples:**
- "Equipment Design Review" template with 10-step checklist
- "CAD Drawing Update" template with standard subtasks
- "Project Kickoff" template with team assignments

**Benefits:**
- Save 5-10 minutes per similar task
- Ensure consistency
- Capture best practices
- Onboard new team members faster

### 5. Advanced Search & Filtering 🔍
**Location:** Enhanced search box in main window

**How to use:**
- **Full-text search:** Type keywords, searches all fields (title, description, notes, links)
- **Multi-field filters:** Click "Advanced Search" button
  - Date ranges: Created between, Due between, Completed between
  - Multiple users: Select multiple assignees
  - Multiple categories: Select multiple categories
  - Tags/Keywords: AND/OR logic
  - Priority levels: Any combination of high/medium/low
- **Saved filters:** Save common searches as presets
  - "My Critical Tasks" = High priority + Assigned to me + Not completed
  - "Overdue Last Month" = Due date < 30 days ago + Not completed
- **Quick filters:** Keyboard shortcuts
  - `Ctrl+F` = Focus search box
  - `F3` = Open advanced search dialog
  - `Ctrl+1/2/3` = Load saved filter #1/2/3

**Database:**
- Table: `SavedFilters` (stores user presets as JSON)
- Indexes: Optimized for fast text search on title/description
- Views: N/A (dynamic queries)

**Search modes:**
- **Simple:** Type in search box → Searches title, description, vessel, drawing no.
- **Advanced:** Click button → Complex multi-field criteria
- **Saved:** Load preset → One-click common searches

**Benefits:**
- Find tasks faster (seconds vs. minutes)
- Complex queries without SQL knowledge
- Save frequently used searches
- Better organization

## Keyboard Shortcuts (New)
| Shortcut | Action |
|----------|--------|
| `Ctrl+N` | New Task |
| `Ctrl+E` | Edit Selected Task |
| `Delete` | Delete Selected Tasks |
| `F5` | Refresh Task List |
| `Ctrl+F` | Focus Search Box |
| `F3` | Advanced Search Dialog |
| `Ctrl+T` | Start/Stop Timer |
| `Ctrl+D` | Open Dashboard |
| `Ctrl+1/2/3` | Load Saved Filter 1/2/3 |

## User Interface Changes

### Main Window
**New buttons added:**
- "Start Timer" (in context menu when task selected)
- "Dashboard" (in main menu or toolbar)
- "Templates" (in main menu)
- "Advanced Search" (next to search box)

**New columns in task list:**
- "Progress" (for subtask completion %)
- "Time Logged" (total hours tracked)

**Enhanced status bar:**
- Shows active timer if running
- Quick stats: "5 tasks due today | 2 overdue | Timer: 1h 23m"

### Task Detail Dialog
**New sections:**
- "Time Tracking" tab
  - Start/Stop timer button
  - Time entry history (manual edits allowed)
  - Total time logged vs. estimated
- "Subtasks" tab
  - Add/edit/delete checklist items
  - Mark complete with checkboxes
  - Completion % indicator
- "Template" section (when creating task)
  - "Create from template" dropdown
  - Auto-fills all fields from template

### Dashboard Window
**Sections:**
1. **Summary Cards** (top row)
   - Total Active Tasks
   - Completed This Month
   - Overdue Count
   - Avg Completion Days

2. **Charts** (middle area)
   - Pie Chart: Tasks by Status
   - Line Chart: Completion Trend (last 30 days)
   - Bar Chart: Workload by User
   - Heatmap: Tasks by Day of Week

3. **Quick Lists** (bottom)
   - Top 5 Overdue Tasks
   - Tasks Due Today
   - Recently Completed

4. **Time Analysis** (right panel)
   - Time Logged vs. Estimated (gauge chart)
   - Top Time Consumers (this week)
   - User Time Distribution

## Configuration (config.ini)
Add these sections if not present:

```ini
[Features]
enable_time_tracking = true
enable_subtasks = true
enable_dashboard = true
enable_templates = true
enable_advanced_search = true
auto_refresh_dashboard_minutes = 5

[TimeTracking]
auto_stop_timer_after_hours = 8
warn_long_sessions = true
long_session_threshold_hours = 4

[Dashboard]
default_date_range = this_month
cache_refresh_minutes = 15
```

## Permissions Required
Ensure your database users have permissions on new tables:

```sql
-- Run this as DBA
USE TaskManagerDB;
GRANT SELECT, INSERT, UPDATE, DELETE ON TimeEntries TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT, INSERT, UPDATE, DELETE ON Subtasks TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT, INSERT, UPDATE, DELETE ON TaskTemplates TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT, INSERT, UPDATE, DELETE ON TemplateSubtasks TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT ON TaskAnalytics TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT, INSERT, UPDATE, DELETE ON SavedFilters TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT ON vw_TaskSummary TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT ON vw_TaskSummaryByUser TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT ON vw_TimeTrackingSummary TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
GRANT SELECT ON vw_SubtaskProgress TO TaskUser1, TaskUser2, TaskUser3, TaskUser4, TaskUser5, TaskUser6, TaskUser7;
```

## Testing Checklist
After implementation, test each feature:

### Time Tracking
- [ ] Start timer on a task
- [ ] Timer updates every minute in UI
- [ ] Stop timer → Entry saved to database
- [ ] Manually add time entry (past date/time)
- [ ] Edit existing time entry
- [ ] Delete time entry
- [ ] View time history for task
- [ ] Total time shown in task list

### Subtasks
- [ ] Add subtask to task
- [ ] Mark subtask complete
- [ ] Reorder subtasks (drag/drop)
- [ ] Delete subtask
- [ ] Completion % updates in main list
- [ ] Subtasks deleted when parent task deleted
- [ ] Assign subtask to different user

### Dashboard
- [ ] Open dashboard from main menu
- [ ] All charts load with data
- [ ] Filter by date range updates charts
- [ ] Export to PDF works
- [ ] Export to Excel works
- [ ] Dashboard refreshes automatically
- [ ] Overdue task list shows correct tasks

### Templates
- [ ] Create new template
- [ ] Add subtasks to template
- [ ] Create task from template (all fields populated)
- [ ] Edit existing template
- [ ] Delete template
- [ ] Share template (public checkbox)
- [ ] Use another user's public template
- [ ] Template usage count increments

### Advanced Search
- [ ] Simple text search works
- [ ] Advanced search dialog opens
- [ ] Multi-field search returns correct results
- [ ] Date range filter works
- [ ] Save custom filter
- [ ] Load saved filter
- [ ] Delete saved filter
- [ ] Search performance acceptable (<2 seconds)

## Troubleshooting

### Issue: "TimeEntries table not found"
**Solution:** Run `sql/create_enhanced_features.sql` script in SSMS

### Issue: Timer doesn't start
**Solution:** 
- Check database connection
- Verify permissions on TimeEntries table
- Check task_manager.log for errors

### Issue: Dashboard shows no data
**Solution:**
- Ensure you have completed tasks in the system
- Check if views exist: `SELECT * FROM vw_TaskSummary`
- Refresh cache: Run the analytics update stored procedure

### Issue: Subtasks don't save
**Solution:**
- Verify Subtasks table exists
- Check for foreign key constraint errors (task_id must exist)
- Review log file for SQL errors

### Issue: Search is slow
**Solution:**
- Rebuild indexes: `ALTER INDEX ALL ON Tasks REBUILD`
- Ensure `IX_Tasks_Title_Description` index exists
- Reduce search result limit in settings

## Performance Considerations

### Database
- Indexes added for fast queries (see schema script)
- Views use appropriate JOINs and WHERE clauses
- TaskAnalytics cache updated daily (reduces query load)

### Application
- Time tracking timer runs in background thread
- Dashboard caches data for 5 minutes
- Search uses parameterized queries (prevents SQL injection)
- Lazy loading: Features only load when accessed

### Network
- Batch database operations where possible
- Async loading for dashboard charts
- Optimistic UI updates (update UI before DB confirms)

## Security Notes

- All database queries use parameterized SQL (prevents injection)
- User permissions enforced at database level
- Saved filters sanitized before saving JSON
- Time entries cannot be modified by other users
- Templates marked public/private per user

## Future Enhancements (Not in v2.0)

- Task dependencies (blockers)
- Gantt chart view
- Email notifications for overdue tasks
- Mobile app integration
- Task comments/activity log
- File attachments
- Recurring tasks
- Custom fields per category
- API for external integrations

## Support
If you encounter issues:
1. Check `task_manager.log` file in application directory
2. Verify database schema is up to date
3. Confirm user permissions
4. Contact IT support with log file attached

---

**Version:** 2.0  
**Last Updated:** January 22, 2026  
**Author:** GitHub Copilot (Claude Sonnet 4.5)
