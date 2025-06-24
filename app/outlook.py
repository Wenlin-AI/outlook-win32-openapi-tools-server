"""Thin wrapper around win32com Outlook API."""

from __future__ import annotations

import os
import sys
import logging
import platform
from datetime import datetime
from typing import Any, Dict, List, Optional
from dotenv import load_dotenv, find_dotenv

# Configure basic logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

try:
    import win32com.client  # type: ignore
    WIN32_AVAILABLE = True
    # Use hardcoded constants for olFolderTasks (13) and olTaskComplete (2)
    # since constants might not be available depending on how win32com is imported
    OUTLOOK_TASK_FOLDER = 13
    OUTLOOK_TASK_COMPLETE = 2
except Exception:  # noqa: BLE001
    win32com = None  # type: ignore
    WIN32_AVAILABLE = False
    OUTLOOK_TASK_FOLDER = 13
    OUTLOOK_TASK_COMPLETE = 2


class OutlookError(Exception):
    """Custom exception for Outlook errors."""


class OutlookTasks:
    """Helper for interacting with Outlook Tasks via COM."""
    
    def __init__(self, folder_path: Optional[str] = None, profile: Optional[str] = None) -> None:
        """Initialize the Outlook Tasks helper with optional folder path and profile."""
        logger = logging.getLogger(__name__)
        
        if platform.system() != "Windows" or not WIN32_AVAILABLE:
            raise OutlookError("pywin32 not available or platform not Windows")

        logger.debug(f"Initializing OutlookTasks with folder_path={folder_path}, profile={profile}")
        
        self.outlook = self._connect_outlook(profile)
        self.namespace = self.outlook.GetNamespace("MAPI")
        
        # Simplified approach: Always use the default Tasks folder
        # unless folder_path is "custom" and then use path resolution
        if folder_path and folder_path.lower() != "default":
            try:
                logger.debug(f"Attempting to find folder from path: {folder_path}")
                entry_id = self._find_folder_id_by_path(folder_path)
                self.tasks_folder = self.namespace.GetFolderFromID(entry_id)
                logger.debug(f"Using custom tasks folder: {self.tasks_folder.Name}")
            except Exception as e:
                logger.error(f"Failed to get custom Tasks folder: {str(e)}")
                raise OutlookError(f"Failed to get custom Tasks folder: {str(e)}")
        else:
            logger.debug("Using default Tasks folder")
            self.tasks_folder = self.namespace.GetDefaultFolder(OUTLOOK_TASK_FOLDER)

        logger.info(f"Using tasks folder: {self.tasks_folder.Name}")

        # Mapping of Outlook EntryID -> sequential task ID. IDs are assigned
        # per-process and rebuilt from Outlook items on demand.
        self._id_map: Dict[str, int] = {}
        self._reverse_map: Dict[int, str] = {}
        self._next_id = 1
    @staticmethod
    def _connect_outlook(profile: Optional[str] = None):
        """Connect to Outlook and initialize COM for the current thread."""
        if not WIN32_AVAILABLE or win32com is None:
            raise OutlookError("pywin32 not available or platform not Windows")
        
        import logging
        logger = logging.getLogger(__name__)
        
        try:
            # Import required modules inside the function
            import pythoncom
            logger.debug("Initializing COM for the current thread...")
            
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            logger.debug("COM initialized successfully")
            
            # Now connect to Outlook
            logger.debug("Attempting to connect to Outlook...")
            app = win32com.client.Dispatch("Outlook.Application")
            logger.debug("Connected to Outlook using Dispatch")
            return app
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {str(e)}", exc_info=True)
            raise OutlookError(f"Failed to connect to Outlook: {str(e)}")
    
    def _find_folder_id_by_path(self, path: str) -> str:
        """Resolve a folder path like '\\Mailbox\\Tasks' to EntryID."""
        import logging
        logger = logging.getLogger(__name__)
        
        # Clean up the path by removing empty parts
        parts = [p for p in path.split("\\") if p]
        logger.debug(f"Folder path parts: {parts}")
        
        if not parts:
            raise OutlookError("Invalid folder path: empty path")
        
        # Start with the root folders collection
        folders = self.namespace.Folders
        current = None
        
        # First level is the mailbox/store name
        for i in range(folders.Count):
            folder = folders.Item(i+1)  # 1-based index
            logger.debug(f"Checking root folder: {folder.Name}")
            if folder.Name == parts[0]:
                current = folder
                logger.debug(f"Found root folder: {folder.Name}")
                break
        
        if current is None:
            raise OutlookError(f"Root folder not found: {parts[0]}")
        
        # Navigate through the remaining path
        for name in parts[1:]:
            found = False
            try:
                sub_folders = current.Folders
                for i in range(sub_folders.Count):
                    folder = sub_folders.Item(i+1)  # 1-based index
                    logger.debug(f"Checking subfolder: {folder.Name}")
                    if folder.Name == name:
                        current = folder
                        found = True
                        logger.debug(f"Found subfolder: {folder.Name}")
                        break
                
                if not found:
                    raise OutlookError(f"Subfolder not found: {name}")
                    
            except Exception as e:
                raise OutlookError(f"Error accessing folder {name}: {str(e)}")
        
        return current.EntryID

    # internal helpers -----------------------------------------------------

    def _ensure_task_id(self, entry_id: str) -> int:
        """Return stable task ID for an EntryID, assigning a new one if needed."""
        if entry_id not in self._id_map:
            self._id_map[entry_id] = self._next_id
            self._reverse_map[self._next_id] = entry_id
            self._next_id += 1
        return self._id_map[entry_id]

    def _build_id_map(self) -> None:
        """Enumerate tasks and ensure each has a task ID."""
        items = self.tasks_folder.Items
        items = items.Restrict("[MessageClass] = 'IPM.Task'")
        for task in items:
            self._ensure_task_id(task.EntryID)

    def _entry_from_task_id(self, task_id: int) -> str:
        try:
            return self._reverse_map[task_id]
        except KeyError:
            # Rebuild mapping in case the process was restarted or mapping was
            # not yet populated. This allows clients to reference tasks without
            # listing them first.
            self._build_id_map()
            try:
                return self._reverse_map[task_id]
            except KeyError as exc:  # pragma: no cover - thin validation
                raise OutlookError(f"Task ID {task_id} not found") from exc

    # Task operations -----------------------------------------------------

    def list_incomplete_tasks(self) -> List[Dict[str, Any]]:
        """List all incomplete tasks in the tasks folder."""
        # Refresh mapping so every visible task has a numeric ID
        self._build_id_map()

        items = self.tasks_folder.Items
        items = items.Restrict("[MessageClass] = 'IPM.Task'")
        incomplete = items.Restrict("[Complete] = 0")

        results = []
        for task in incomplete:
            task_id = self._ensure_task_id(task.EntryID)
            results.append(
                {
                    "taskId": task_id,
                    "entryId": task.EntryID,
                    "subject": task.Subject,
                    "dueDate": getattr(task, "DueDate", None),
                    "status": task.Status,
                    "body": getattr(task, "Body", None),
                }
            )
        return results

    def add_task(
        self, subject: str, due_date: Optional[datetime] = None, body: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a new task to Outlook Tasks folder."""
        task = self.tasks_folder.Items.Add("IPM.Task")
        task.Subject = subject
        
        if due_date:
            task.DueDate = due_date
            
        if body:
            task.Body = body
            
        task.Save()
        task_id = self._ensure_task_id(task.EntryID)
        return {"entryId": task.EntryID, "taskId": task_id}

    def complete_task(self, task_id: int) -> None:
        entry_id = self._entry_from_task_id(task_id)
        task = self.namespace.GetItemFromID(entry_id)
        task.Status = OUTLOOK_TASK_COMPLETE  # olTaskComplete = 2
        task.Save()

    def delete_task(self, task_id: int) -> None:
        entry_id = self._entry_from_task_id(task_id)
        task = self.namespace.GetItemFromID(entry_id)
        task.Delete()
        self._id_map.pop(entry_id, None)
        self._reverse_map.pop(task_id, None)


# Convenience factory ------------------------------------------------------

def get_tasks_client() -> OutlookTasks:
    """Factory function to create an OutlookTasks client using environment variables."""
    import logging
    logger = logging.getLogger(__name__)
    
    # Make sure we load from .env file before creating the client
    dotenv_path = find_dotenv(usecwd=True)
    if dotenv_path:
        logger.debug(f"Loading environment from: {dotenv_path}")
        load_dotenv(dotenv_path)
    else:
        logger.debug("No .env file found. Using default environment variables.")
        
    folder = os.getenv("OUTLOOK_TASKS_FOLDER")
    profile = os.getenv("OUTLOOK_PROFILE")
    
    logger.debug(f"Creating OutlookTasks with folder={folder}, profile={profile}")
    
    try:
        client = OutlookTasks(folder, profile)
        logger.debug("OutlookTasks client created successfully")
        return client
    except Exception as e:
        logger.error(f"Failed to create OutlookTasks client: {str(e)}", exc_info=True)
        raise

def get_default_task_folders() -> List[Dict[str, str]]:
    """Get only the default task folders from Outlook.
    
    This function connects to Outlook and retrieves only the Tasks folder
    for each mail account. Returns a list of dictionaries with folder name 
    and path that can be used in the OUTLOOK_TASKS_FOLDER environment variable.
    """
    import logging
    logger = logging.getLogger(__name__)
    
    if platform.system() != "Windows" or not WIN32_AVAILABLE:
        raise OutlookError("pywin32 not available or platform not Windows")
    
    folders = []
    
    # Initialize COM for this thread
    try:
        import pythoncom
        pythoncom.CoInitialize()
        
        # Create our own connection to Outlook
        if win32com is not None:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Get only the Tasks folder (constant value 13)
            tasks_folder_const = OUTLOOK_TASK_FOLDER  # olFolderTasks = 13
            
            # Try to get the default Tasks folder from the primary account
            try:
                default_tasks = namespace.GetDefaultFolder(tasks_folder_const)
                folders.append({
                    "name": "Default Tasks",
                    "path": default_tasks.FolderPath
                })
                logger.info(f"Found default Tasks folder: {default_tasks.FolderPath}")
            except Exception as e:
                logger.error(f"Error accessing default Tasks folder: {e}")
                
            # Try the specific path
            # Try a specific path from configuration or environment
            try:
                specific_path = os.getenv("OUTLOOK_TASKS_FOLDER", None)
                if not specific_path or specific_path.lower() == "default":
                    logger.info("No specific tasks folder path provided in configuration; skipping.")
                    raise Exception("No specific tasks folder path provided.")
                logger.info(f"Looking for specific tasks folder: {specific_path}")
                
                # Clean up the path by removing empty parts
                parts = [p for p in specific_path.split("\\") if p]
                
                # Start with the root folders collection
                root_folders = namespace.Folders
                current = None
                # First level is the mailbox/store name
                for i in range(root_folders.Count):
                    folder = root_folders.Item(i+1)  # 1-based index
                    logger.debug(f"Checking root folder: {folder.Name}")
                    if folder.Name == parts[0]:
                        current = folder
                        logger.debug(f"Found root folder: {folder.Name}")
                        break
                
                if current is not None:
                    # Navigate through the remaining path
                    for name in parts[1:]:
                        found = False
                        sub_folders = current.Folders
                        for i in range(sub_folders.Count):
                            folder = sub_folders.Item(i+1)  # 1-based index
                            logger.debug(f"Checking subfolder: {folder.Name}")
                            if folder.Name == name:
                                current = folder
                                found = True
                                logger.debug(f"Found subfolder: {folder.Name}")
                                break
                        
                        if not found:
                            logger.warning(f"Subfolder not found: {name}")
                            raise Exception(f"Subfolder not found: {name}")
                    
                    # If we got here, we found the folder
                    folders.append({
                        "name": "Known Tasks Folder",
                        "path": specific_path
                    })
                    logger.info(f"Found specific Tasks folder: {specific_path}")
                else:
                    logger.warning(f"Root folder not found: {parts[0]}")
            except Exception as e:
                logger.error(f"Error accessing specific Tasks folder: {e}")
    except Exception as e:
        logger.error(f"Failed to connect to Outlook: {str(e)}")
        raise OutlookError(f"Failed to connect to Outlook: {str(e)}")
    
    return folders


def load_env_config():
    """Load environment variables from .env file."""
    import logging
    logger = logging.getLogger(__name__)
    # Look for .env in current directory and parent directories
    dotenv_path = find_dotenv(usecwd=True)
    if dotenv_path:
        logger.info(f"Loading environment from: {dotenv_path}")
        load_dotenv(dotenv_path)
    else:
        logger.info("No .env file found. Using default environment variables.")
    
    # Log current configuration (without passwords or secrets)
    logger.info(f"OUTLOOK_TASKS_FOLDER: {os.getenv('OUTLOOK_TASKS_FOLDER', 'Not set')}")
    logger.info(f"OUTLOOK_PROFILE: {os.getenv('OUTLOOK_PROFILE', 'Not set')}")

if __name__ == "__main__":
    if platform.system() != "Windows" or not WIN32_AVAILABLE:
        print("This script requires Windows and the pywin32 package.")
        print("Please install it with: pip install pywin32")
        sys.exit(1)
    
    # Load environment variables from .env file
    print("\nLoading environment configuration...")
    load_env_config()

    outlook = OutlookTasks(folder_path=os.getenv("OUTLOOK_TASKS_FOLDER"),
                           profile=os.getenv("OUTLOOK_PROFILE"))
    print("\nListing incomplete tasks:")
    tasks = outlook.list_incomplete_tasks()
    for task in tasks:
        print(f"- {task['subject']} (Due: {task['dueDate']}, Status: {task['status']})")
    if not tasks:
        print("No incomplete tasks found.")
