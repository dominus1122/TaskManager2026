"""
Task Manager Enhanced Features Module
=====================================
This module provides 5 major feature enhancements:
1. Task Time Tracking
2. Subtasks/Checklists
3. Dashboard/Analytics
4. Task Templates
5. Advanced Search & Filtering

Version: 2.0
Date: January 22, 2026
"""

import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import Dict, List, Optional, Any, Tuple
import logging
import json
import threading
import time as time_module

# ============================================
# FEATURE 1: TASK TIME TRACKING
# ============================================

class TimeTrackingManager:
    """Manages time tracking for tasks."""
    
    def __init__(self, task_manager):
        """Initialize time tracking manager.
        
        Args:
            task_manager: Reference to main TaskManager instance
        """
        self.task_manager = task_manager
        self.active_timers = {}  # {task_id: start_datetime}
        self.timer_threads = {}  # {task_id: thread_obj}
        
    def start_timer(self, task_id: int, user_name: str) -> bool:
        """Start timer for a task."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            logging.error("Cannot start timer: No database connection")
            return False
            
        if task_id in self.active_timers:
            logging.warning(f"Timer already running for task {task_id}")
            return False
        
        try:
            start_time = datetime.datetime.now()
            
            # Insert time entry with null end_time (timer running)
            self.task_manager.cursor.execute("""
                INSERT INTO TimeEntries (task_id, user_name, start_time, entry_type)
                VALUES (%s, %s, %s, 'timer')
            """, (task_id, user_name, start_time))
            
            # Get the inserted time_entry_id
            self.task_manager.cursor.execute("SELECT @@IDENTITY")
            time_entry_id = self.task_manager.cursor.fetchone()[0]
            
            self.active_timers[task_id] = {
                'start_time': start_time,
                'time_entry_id': time_entry_id,
                'user_name': user_name
            }
            
            logging.info(f"Timer started for task {task_id} by {user_name}")
            return True
            
        except Exception as e:
            logging.error(f"Error starting timer for task {task_id}: {e}", exc_info=True)
            return False
    
    def stop_timer(self, task_id: int) -> Optional[int]:
        """Stop timer for a task and return duration in minutes."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            logging.error("Cannot stop timer: No database connection")
            return None
            
        if task_id not in self.active_timers:
            logging.warning(f"No active timer for task {task_id}")
            return None
        
        try:
            timer_info = self.active_timers[task_id]
            end_time = datetime.datetime.now()
            start_time = timer_info['start_time']
            time_entry_id = timer_info['time_entry_id']
            
            # Calculate duration
            duration_seconds = (end_time - start_time).total_seconds()
            duration_minutes = int(duration_seconds / 60)
            
            # Update time entry
            self.task_manager.cursor.execute("""
                UPDATE TimeEntries
                SET end_time = %s, duration_minutes = %s
                WHERE time_entry_id = %s
            """, (end_time, duration_minutes, time_entry_id))
            
            del self.active_timers[task_id]
            
            logging.info(f"Timer stopped for task {task_id}. Duration: {duration_minutes} minutes")
            return duration_minutes
            
        except Exception as e:
            logging.error(f"Error stopping timer for task {task_id}: {e}", exc_info=True)
            return None
    
    def is_timer_running(self, task_id: int) -> bool:
        """Check if timer is running for a task."""
        return task_id in self.active_timers
    
    def get_timer_duration(self, task_id: int) -> Optional[int]:
        """Get current duration of running timer in minutes."""
        if task_id not in self.active_timers:
            return None
        
        start_time = self.active_timers[task_id]['start_time']
        now = datetime.datetime.now()
        duration_seconds = (now - start_time).total_seconds()
        return int(duration_seconds / 60)
    
    def add_manual_time_entry(self, task_id: int, user_name: str, 
                             start_time: datetime.datetime, 
                             duration_minutes: int,
                             description: str = "") -> bool:
        """Add a manual time entry."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return False
        
        try:
            end_time = start_time + datetime.timedelta(minutes=duration_minutes)
            
            self.task_manager.cursor.execute("""
                INSERT INTO TimeEntries 
                (task_id, user_name, start_time, end_time, duration_minutes, description, entry_type)
                VALUES (%s, %s, %s, %s, %s, %s, 'manual')
            """, (task_id, user_name, start_time, end_time, duration_minutes, description))
            
            logging.info(f"Manual time entry added for task {task_id}: {duration_minutes} minutes")
            return True
            
        except Exception as e:
            logging.error(f"Error adding manual time entry: {e}", exc_info=True)
            return False
    
    def get_time_entries(self, task_id: int) -> List[Dict[str, Any]]:
        """Get all time entries for a task."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            self.task_manager.cursor.execute("""
                SELECT time_entry_id, user_name, start_time, end_time, 
                       duration_minutes, description, entry_type, created_date
                FROM TimeEntries
                WHERE task_id = %s AND is_deleted = 0
                ORDER BY start_time DESC
            """, (task_id,))
            
            entries = []
            for row in self.task_manager.cursor.fetchall():
                entries.append({
                    'time_entry_id': row[0],
                    'user_name': row[1],
                    'start_time': row[2],
                    'end_time': row[3],
                    'duration_minutes': row[4],
                    'description': row[5],
                    'entry_type': row[6],
                    'created_date': row[7]
                })
            
            return entries
            
        except Exception as e:
            logging.error(f"Error fetching time entries: {e}", exc_info=True)
            return []
    
    def get_total_time_logged(self, task_id: int) -> int:
        """Get total time logged for a task in minutes."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return 0
        
        try:
            self.task_manager.cursor.execute("""
                SELECT COALESCE(SUM(duration_minutes), 0)
                FROM TimeEntries
                WHERE task_id = %s AND is_deleted = 0
            """, (task_id,))
            
            return self.task_manager.cursor.fetchone()[0]
            
        except Exception as e:
            logging.error(f"Error calculating total time: {e}", exc_info=True)
            return 0


# ============================================
# FEATURE 2: SUBTASKS/CHECKLISTS
# ============================================

class SubtaskManager:
    """Manages subtasks/checklists for tasks."""
    
    def __init__(self, task_manager):
        """Initialize subtask manager."""
        self.task_manager = task_manager
    
    def add_subtask(self, task_id: int, title: str, description: str = "",
                   assigned_to: str = None, due_date: str = None,
                   estimated_minutes: int = None) -> Optional[int]:
        """Add a subtask to a task."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return None
        
        try:
            # Get next sort_order
            self.task_manager.cursor.execute("""
                SELECT COALESCE(MAX(sort_order), -1) + 1
                FROM Subtasks
                WHERE task_id = %s AND is_deleted = 0
            """, (task_id,))
            sort_order = self.task_manager.cursor.fetchone()[0]
            
            current_user = self.task_manager._get_current_user()
            
            self.task_manager.cursor.execute("""
                INSERT INTO Subtasks 
                (task_id, title, description, sort_order, assigned_to, due_date, 
                 estimated_minutes, created_by)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (task_id, title, description, sort_order, assigned_to, due_date,
                  estimated_minutes, current_user))
            
            self.task_manager.cursor.execute("SELECT @@IDENTITY")
            subtask_id = self.task_manager.cursor.fetchone()[0]
            
            logging.info(f"Subtask {subtask_id} added to task {task_id}")
            return subtask_id
            
        except Exception as e:
            logging.error(f"Error adding subtask: {e}", exc_info=True)
            return None
    
    def get_subtasks(self, task_id: int) -> List[Dict[str, Any]]:
        """Get all subtasks for a task."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            self.task_manager.cursor.execute("""
                SELECT subtask_id, title, description, completed, sort_order,
                       estimated_minutes, actual_minutes, assigned_to, due_date,
                       completed_date, created_date, created_by
                FROM Subtasks
                WHERE task_id = %s AND is_deleted = 0
                ORDER BY sort_order, subtask_id
            """, (task_id,))
            
            subtasks = []
            for row in self.task_manager.cursor.fetchall():
                subtasks.append({
                    'subtask_id': row[0],
                    'title': row[1],
                    'description': row[2],
                    'completed': bool(row[3]),
                    'sort_order': row[4],
                    'estimated_minutes': row[5],
                    'actual_minutes': row[6],
                    'assigned_to': row[7],
                    'due_date': row[8].strftime('%Y-%m-%d') if row[8] else None,
                    'completed_date': row[9],
                    'created_date': row[10],
                    'created_by': row[11]
                })
            
            return subtasks
            
        except Exception as e:
            logging.error(f"Error fetching subtasks: {e}", exc_info=True)
            return []
    
    def toggle_subtask(self, subtask_id: int) -> bool:
        """Toggle subtask completion status."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return False
        
        try:
            # Get current status
            self.task_manager.cursor.execute("""
                SELECT completed FROM Subtasks WHERE subtask_id = %s
            """, (subtask_id,))
            current = self.task_manager.cursor.fetchone()[0]
            
            new_status = not bool(current)
            completed_date = datetime.datetime.now() if new_status else None
            
            self.task_manager.cursor.execute("""
                UPDATE Subtasks
                SET completed = %s, completed_date = %s
                WHERE subtask_id = %s
            """, (new_status, completed_date, subtask_id))
            
            return True
            
        except Exception as e:
            logging.error(f"Error toggling subtask: {e}", exc_info=True)
            return False
    
    def delete_subtask(self, subtask_id: int) -> bool:
        """Soft delete a subtask."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return False
        
        try:
            self.task_manager.cursor.execute("""
                UPDATE Subtasks SET is_deleted = 1 WHERE subtask_id = %s
            """, (subtask_id,))
            return True
        except Exception as e:
            logging.error(f"Error deleting subtask: {e}", exc_info=True)
            return False
    
    def get_completion_percentage(self, task_id: int) -> int:
        """Get completion percentage for a task's subtasks."""
        subtasks = self.get_subtasks(task_id)
        if not subtasks:
            return 0
        
        completed = sum(1 for s in subtasks if s['completed'])
        return int((completed / len(subtasks)) * 100)


# ============================================
# FEATURE 3: TASK TEMPLATES
# ============================================

class TemplateManager:
    """Manages task templates."""
    
    def __init__(self, task_manager):
        """Initialize template manager."""
        self.task_manager = task_manager
    
    def create_template(self, template_name: str, description: str = "",
                       category: str = None, priority: str = "medium",
                       estimated_mhr: int = None, default_duration_days: int = None,
                       default_main_staff: str = None, default_assigned_to: str = None,
                       is_public: bool = True) -> Optional[int]:
        """Create a new task template."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return None
        
        try:
            created_by = self.task_manager._get_current_user()
            
            self.task_manager.cursor.execute("""
                INSERT INTO TaskTemplates 
                (template_name, description, category, priority, estimated_mhr,
                 default_duration_days, default_main_staff, default_assigned_to,
                 is_public, created_by)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (template_name, description, category, priority, estimated_mhr,
                  default_duration_days, default_main_staff, default_assigned_to,
                  is_public, created_by))
            
            self.task_manager.cursor.execute("SELECT @@IDENTITY")
            template_id = self.task_manager.cursor.fetchone()[0]
            
            logging.info(f"Template '{template_name}' created with ID {template_id}")
            return template_id
            
        except Exception as e:
            logging.error(f"Error creating template: {e}", exc_info=True)
            return None
    
    def get_templates(self, user_name: str = None, include_public: bool = True) -> List[Dict[str, Any]]:
        """Get all available templates for a user."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            if user_name and include_public:
                query = """
                    SELECT template_id, template_name, description, category, priority,
                           estimated_mhr, default_duration_days, default_main_staff,
                           default_assigned_to, is_public, created_by, usage_count
                    FROM TaskTemplates
                    WHERE is_deleted = 0 AND (is_public = 1 OR created_by = %s)
                    ORDER BY usage_count DESC, template_name
                """
                self.task_manager.cursor.execute(query, (user_name,))
            else:
                query = """
                    SELECT template_id, template_name, description, category, priority,
                           estimated_mhr, default_duration_days, default_main_staff,
                           default_assigned_to, is_public, created_by, usage_count
                    FROM TaskTemplates
                    WHERE is_deleted = 0
                    ORDER BY usage_count DESC, template_name
                """
                self.task_manager.cursor.execute(query)
            
            templates = []
            for row in self.task_manager.cursor.fetchall():
                templates.append({
                    'template_id': row[0],
                    'template_name': row[1],
                    'description': row[2],
                    'category': row[3],
                    'priority': row[4],
                    'estimated_mhr': row[5],
                    'default_duration_days': row[6],
                    'default_main_staff': row[7],
                    'default_assigned_to': row[8],
                    'is_public': bool(row[9]),
                    'created_by': row[10],
                    'usage_count': row[11]
                })
            
            return templates
            
        except Exception as e:
            logging.error(f"Error fetching templates: {e}", exc_info=True)
            return []
    
    def get_template(self, template_id: int) -> Optional[Dict[str, Any]]:
        """Get a specific template by ID."""
        templates = self.get_templates()
        for template in templates:
            if template['template_id'] == template_id:
                # Also fetch template subtasks
                template['subtasks'] = self.get_template_subtasks(template_id)
                return template
        return None
    
    def add_template_subtask(self, template_id: int, title: str, 
                            description: str = "", estimated_minutes: int = None) -> Optional[int]:
        """Add a subtask to a template."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return None
        
        try:
            # Get next sort_order
            self.task_manager.cursor.execute("""
                SELECT COALESCE(MAX(sort_order), -1) + 1
                FROM TemplateSubtasks
                WHERE template_id = %s AND is_deleted = 0
            """, (template_id,))
            sort_order = self.task_manager.cursor.fetchone()[0]
            
            self.task_manager.cursor.execute("""
                INSERT INTO TemplateSubtasks (template_id, title, description, sort_order, estimated_minutes)
                VALUES (%s, %s, %s, %s, %s)
            """, (template_id, title, description, sort_order, estimated_minutes))
            
            self.task_manager.cursor.execute("SELECT @@IDENTITY")
            return self.task_manager.cursor.fetchone()[0]
            
        except Exception as e:
            logging.error(f"Error adding template subtask: {e}", exc_info=True)
            return None
    
    def get_template_subtasks(self, template_id: int) -> List[Dict[str, Any]]:
        """Get all subtasks for a template."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            self.task_manager.cursor.execute("""
                SELECT template_subtask_id, title, description, sort_order, estimated_minutes
                FROM TemplateSubtasks
                WHERE template_id = %s AND is_deleted = 0
                ORDER BY sort_order
            """, (template_id,))
            
            subtasks = []
            for row in self.task_manager.cursor.fetchall():
                subtasks.append({
                    'template_subtask_id': row[0],
                    'title': row[1],
                    'description': row[2],
                    'sort_order': row[3],
                    'estimated_minutes': row[4]
                })
            
            return subtasks
            
        except Exception as e:
            logging.error(f"Error fetching template subtasks: {e}", exc_info=True)
            return []
    
    def increment_usage_count(self, template_id: int):
        """Increment usage counter for a template."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return
        
        try:
            self.task_manager.cursor.execute("""
                UPDATE TaskTemplates
                SET usage_count = usage_count + 1, modified_date = GETDATE()
                WHERE template_id = %s
            """, (template_id,))
        except Exception as e:
            logging.error(f"Error incrementing usage count: {e}")


# ============================================
# FEATURE 4: ADVANCED SEARCH
# ============================================

class AdvancedSearchManager:
    """Manages advanced search and saved filters."""
    
    def __init__(self, task_manager):
        """Initialize search manager."""
        self.task_manager = task_manager
    
    def search_tasks(self, search_params: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Perform advanced search with multiple criteria.
        
        Args:
            search_params: Dictionary with search criteria:
                - text: Full-text search
                - categories: List of categories
                - priorities: List of priorities
                - assigned_to: List of users
                - main_staff: List of main staff
                - date_created_start, date_created_end
                - date_due_start, date_due_end
                - completed: True/False/None (all)
                - overdue_only: True/False
        """
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            # Build dynamic WHERE clause
            conditions = ["deleted = 0"]
            params = []
            
            # Full-text search
            if search_params.get('text'):
                text = f"%{search_params['text']}%"
                conditions.append("""
                    (title LIKE %s OR description LIKE %s OR 
                     applied_vessel LIKE %s OR drawing_no LIKE %s OR 
                     request_no LIKE %s OR link LIKE %s)
                """)
                params.extend([text] * 6)
            
            # Categories
            if search_params.get('categories'):
                placeholders = ', '.join(['%s'] * len(search_params['categories']))
                conditions.append(f"category IN ({placeholders})")
                params.extend(search_params['categories'])
            
            # Priorities
            if search_params.get('priorities'):
                placeholders = ', '.join(['%s'] * len(search_params['priorities']))
                conditions.append(f"priority IN ({placeholders})")
                params.extend(search_params['priorities'])
            
            # Assigned to
            if search_params.get('assigned_to'):
                placeholders = ', '.join(['%s'] * len(search_params['assigned_to']))
                conditions.append(f"assigned_to IN ({placeholders})")
                params.extend(search_params['assigned_to'])
            
            # Main staff
            if search_params.get('main_staff'):
                placeholders = ', '.join(['%s'] * len(search_params['main_staff']))
                conditions.append(f"main_staff IN ({placeholders})")
                params.extend(search_params['main_staff'])
            
            # Date created range
            if search_params.get('date_created_start'):
                conditions.append("created_date >= %s")
                params.append(search_params['date_created_start'])
            if search_params.get('date_created_end'):
                conditions.append("created_date <= %s")
                params.append(search_params['date_created_end'])
            
            # Date due range
            if search_params.get('date_due_start'):
                conditions.append("due_date >= %s")
                params.append(search_params['date_due_start'])
            if search_params.get('date_due_end'):
                conditions.append("due_date <= %s")
                params.append(search_params['date_due_end'])
            
            # Completion status
            if search_params.get('completed') is not None:
                conditions.append("completed = %s")
                params.append(search_params['completed'])
            
            # Overdue only
            if search_params.get('overdue_only'):
                conditions.append("completed = 0 AND due_date < CAST(GETDATE() AS DATE)")
            
            # Build and execute query
            where_clause = " AND ".join(conditions)
            query = f"""
                SELECT id, title, description, due_date, priority, category, completed,
                       main_staff, assigned_to, created_date, last_modified,
                       applied_vessel, rev, drawing_no, link, sdb_link, request_no,
                       requested_date, target_start, target_finish, qtd_mhr, actual_mhr
                FROM Tasks
                WHERE {where_clause}
                ORDER BY priority DESC, due_date ASC
            """
            
            self.task_manager.cursor.execute(query, params)
            
            # Convert to task dictionaries (similar to _load_tasks)
            tasks = []
            for row in self.task_manager.cursor.fetchall():
                task = {
                    'id': row[0],
                    'title': row[1],
                    'description': row[2],
                    'due_date': row[3].strftime('%Y-%m-%d') if row[3] else None,
                    'priority': row[4],
                    'category': row[5],
                    'completed': bool(row[6]),
                    'main_staff': row[7],
                    'assigned_to': row[8],
                    'created_date': row[9].strftime('%Y-%m-%d') if row[9] else None,
                    'last_modified': row[10].strftime('%Y-%m-%d %H:%M:%S') if row[10] else None,
                    'applied_vessel': row[11],
                    'rev': row[12],
                    'drawing_no': row[13],
                    'link': row[14],
                    'sdb_link': row[15],
                    'request_no': row[16],
                    'requested_date': row[17].strftime('%Y-%m-%d') if row[17] else None,
                    'target_start': row[18].strftime('%Y-%m-%d') if row[18] else None,
                    'target_finish': row[19].strftime('%Y-%m-%d') if row[19] else None,
                    'qtd_mhr': row[20],
                    'actual_mhr': row[21],
                    'deleted': False
                }
                tasks.append(task)
            
            logging.info(f"Advanced search returned {len(tasks)} tasks")
            return tasks
            
        except Exception as e:
            logging.error(f"Error in advanced search: {e}", exc_info=True)
            return []
    
    def save_filter(self, user_name: str, filter_name: str, 
                   filter_params: Dict[str, Any], is_public: bool = False) -> bool:
        """Save a search filter for future use."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return False
        
        try:
            filter_json = json.dumps(filter_params)
            
            # Check if filter already exists
            self.task_manager.cursor.execute("""
                SELECT filter_id FROM SavedFilters 
                WHERE user_name = %s AND filter_name = %s
            """, (user_name, filter_name))
            
            if self.task_manager.cursor.fetchone():
                # Update existing
                self.task_manager.cursor.execute("""
                    UPDATE SavedFilters
                    SET filter_json = %s, is_public = %s
                    WHERE user_name = %s AND filter_name = %s
                """, (filter_json, is_public, user_name, filter_name))
            else:
                # Insert new
                self.task_manager.cursor.execute("""
                    INSERT INTO SavedFilters (filter_name, user_name, filter_json, is_public)
                    VALUES (%s, %s, %s, %s)
                """, (filter_name, user_name, filter_json, is_public))
            
            logging.info(f"Filter '{filter_name}' saved for user {user_name}")
            return True
            
        except Exception as e:
            logging.error(f"Error saving filter: {e}", exc_info=True)
            return False
    
    def get_saved_filters(self, user_name: str, include_public: bool = True) -> List[Dict[str, Any]]:
        """Get all saved filters for a user."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            if include_public:
                query = """
                    SELECT filter_id, filter_name, filter_json, is_public, created_date, use_count
                    FROM SavedFilters
                    WHERE user_name = %s OR is_public = 1
                    ORDER BY use_count DESC, filter_name
                """
                self.task_manager.cursor.execute(query, (user_name,))
            else:
                query = """
                    SELECT filter_id, filter_name, filter_json, is_public, created_date, use_count
                    FROM SavedFilters
                    WHERE user_name = %s
                    ORDER BY filter_name
                """
                self.task_manager.cursor.execute(query, (user_name,))
            
            filters = []
            for row in self.task_manager.cursor.fetchall():
                filters.append({
                    'filter_id': row[0],
                    'filter_name': row[1],
                    'filter_params': json.loads(row[2]),
                    'is_public': bool(row[3]),
                    'created_date': row[4],
                    'use_count': row[5]
                })
            
            return filters
            
        except Exception as e:
            logging.error(f"Error fetching saved filters: {e}", exc_info=True)
            return []
    
    def delete_saved_filter(self, filter_id: int) -> bool:
        """Delete a saved filter."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return False
        
        try:
            self.task_manager.cursor.execute("""
                DELETE FROM SavedFilters WHERE filter_id = %s
            """, (filter_id,))
            return True
        except Exception as e:
            logging.error(f"Error deleting filter: {e}", exc_info=True)
            return False
    
    def increment_filter_usage(self, filter_id: int):
        """Increment usage counter for a filter."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return
        
        try:
            self.task_manager.cursor.execute("""
                UPDATE SavedFilters
                SET use_count = use_count + 1, last_used = GETDATE()
                WHERE filter_id = %s
            """, (filter_id,))
        except Exception as e:
            logging.error(f"Error incrementing filter usage: {e}")


# ============================================
# FEATURE 5: DASHBOARD/ANALYTICS
# ============================================

class DashboardManager:
    """Manages dashboard analytics and metrics."""
    
    def __init__(self, task_manager):
        """Initialize dashboard manager."""
        self.task_manager = task_manager
    
    def get_summary_metrics(self) -> Dict[str, Any]:
        """Get overall summary metrics including monthly stats."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return {}
        
        try:
            # 1. Get All-Time Stats from View
            self.task_manager.cursor.execute("""
                SELECT total_tasks, completed_tasks, active_tasks, overdue_tasks,
                       high_priority_active, medium_priority_active, low_priority_active,
                       total_estimated_mhr, total_actual_mhr, avg_completion_days
                FROM vw_TaskSummary
            """)
            
            row = self.task_manager.cursor.fetchone()
            if not row:
                return {}
            
            metrics = {
                'total_tasks': row[0],
                'completed_tasks': row[1],
                'active_tasks': row[2],
                'overdue_tasks': row[3],
                'high_priority_active': row[4],
                'medium_priority_active': row[5],
                'low_priority_active': row[6],
                'total_estimated_mhr': row[7],
                'total_actual_mhr': row[8],
                'avg_completion_days': float(row[9]) if row[9] else 0.0,
                'completion_rate': (row[1] / row[0] * 100) if row[0] > 0 else 0.0
            }

            # 2. Get Monthly Due Stats (Tasks due in current month)
            # This gives a better sense of "Current Progress" than all-time stats
            self.task_manager.cursor.execute("""
                SELECT 
                    COUNT(*) as monthly_due_total,
                    SUM(CASE WHEN completed = 1 THEN 1 ELSE 0 END) as monthly_due_completed
                FROM Tasks 
                WHERE MONTH(due_date) = MONTH(GETDATE()) 
                  AND YEAR(due_date) = YEAR(GETDATE())
                  AND deleted = 0
            """)
            
            monthly_row = self.task_manager.cursor.fetchone()
            if monthly_row:
                monthly_total = monthly_row[0] or 0
                monthly_completed = monthly_row[1] or 0
                metrics['monthly_due_total'] = monthly_total
                metrics['monthly_due_completed'] = monthly_completed
                metrics['monthly_progress'] = (monthly_completed / monthly_total * 100) if monthly_total > 0 else 0.0
            else:
                metrics['monthly_due_total'] = 0
                metrics['monthly_due_completed'] = 0
                metrics['monthly_progress'] = 0.0
            
            return metrics
            
        except Exception as e:
            logging.error(f"Error fetching summary metrics: {e}", exc_info=True)
            return {}
    
    def get_user_workload(self) -> List[Dict[str, Any]]:
        """Get workload distribution by user."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            self.task_manager.cursor.execute("""
                SELECT user_name, total_tasks, completed_tasks, active_tasks,
                       overdue_tasks, total_estimated_mhr, total_actual_mhr
                FROM vw_TaskSummaryByUser
                ORDER BY active_tasks DESC, overdue_tasks DESC
            """)
            
            workload = []
            for row in self.task_manager.cursor.fetchall():
                workload.append({
                    'user_name': row[0],
                    'total_tasks': row[1],
                    'completed_tasks': row[2],
                    'active_tasks': row[3],
                    'overdue_tasks': row[4],
                    'total_estimated_mhr': row[5],
                    'total_actual_mhr': row[6]
                })
            
            return workload
            
        except Exception as e:
            logging.error(f"Error fetching user workload: {e}", exc_info=True)
            return []
    
    def get_completion_trend(self, days: int = 30) -> List[Dict[str, Any]]:
        """Get task completion trend for the last N days."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            # Get tasks completed in the last N days, grouped by date
            self.task_manager.cursor.execute("""
                SELECT CAST(last_modified AS DATE) as completion_date, COUNT(*) as count
                FROM Tasks
                WHERE completed = 1 AND deleted = 0
                  AND last_modified >= DATEADD(day, -%s, GETDATE())
                GROUP BY CAST(last_modified AS DATE)
                ORDER BY completion_date
            """, (days,))
            
            trend = []
            for row in self.task_manager.cursor.fetchall():
                trend.append({
                    'date': row[0].strftime('%Y-%m-%d'),
                    'completed_count': row[1]
                })
            
            return trend
            
        except Exception as e:
            logging.error(f"Error fetching completion trend: {e}", exc_info=True)
            return []

    def get_upcoming_deadlines(self, limit: int = 15) -> List[Dict[str, Any]]:
        """Get upcoming task deadlines (Next 7 Days)."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            # Query for next 7 days (inclusive of today)
            self.task_manager.cursor.execute(f"""
                SELECT TOP {limit} id, title, due_date, priority, assigned_to
                FROM Tasks 
                WHERE completed = 0 AND deleted = 0 
                  AND due_date IS NOT NULL
                  AND CAST(due_date AS DATE) >= CAST(GETDATE() AS DATE)
                  AND CAST(due_date AS DATE) <= DATEADD(day, 7, CAST(GETDATE() AS DATE))
                ORDER BY due_date ASC, 
                         CASE priority 
                             WHEN 'High' THEN 1 
                             WHEN 'Medium' THEN 2 
                             WHEN 'Low' THEN 3 
                             ELSE 4
                         END
            """)
            
            tasks = []
            for row in self.task_manager.cursor.fetchall():
                tasks.append({
                    'task_id': row[0],
                    'title': row[1],
                    'due_date': row[2].strftime('%Y-%m-%d') if row[2] else 'No Date',
                    'priority': row[3],
                    'assigned_to': row[4]
                })
            return tasks
        except Exception as e:
            logging.error(f"Error fetching upcoming deadlines: {e}", exc_info=True)
            # Return error as a special task item to visualize in UI for debugging
            return [{'task_id': -1, 'title': f"Error: {str(e)}", 'due_date': 'Error', 'priority': 'High', 'assigned_to': 'System'}]

    def get_time_tracking_summary(self) -> List[Dict[str, Any]]:
        """Get time tracking summary for active tasks."""
        if not self.task_manager.conn or not self.task_manager.cursor:
            return []
        
        try:
            self.task_manager.cursor.execute("""
                SELECT TOP 20 task_id, title, assigned_to, entry_count,
                       total_minutes_logged, total_hours_logged
                FROM vw_TimeTrackingSummary
                WHERE total_minutes_logged > 0
                ORDER BY total_minutes_logged DESC
            """)
            
            summary = []
            for row in self.task_manager.cursor.fetchall():
                summary.append({
                    'task_id': row[0],
                    'title': row[1],
                    'assigned_to': row[2],
                    'entry_count': row[3],
                    'total_minutes': row[4],
                    'total_hours': float(row[5])
                })
            
            return summary
            
        except Exception as e:
            logging.error(f"Error fetching time tracking summary: {e}", exc_info=True)
            return []


# Export all managers
__all__ = [
    'TimeTrackingManager',
    'SubtaskManager',
    'TemplateManager',
    'AdvancedSearchManager',
    'DashboardManager'
]
