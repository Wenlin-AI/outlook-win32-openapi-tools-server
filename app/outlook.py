"""Thin wrapper around win32com Outlook API."""

from __future__ import annotations

import os
import sys
import platform
from datetime import datetime
from typing import Any, Dict, List, Optional

try:
    import win32com.client  # type: ignore
    from win32com.client import constants  # type: ignore
    WIN32_AVAILABLE = True
except Exception:  # noqa: BLE001
    win32com = None  # type: ignore
    constants = None  # type: ignore
    WIN32_AVAILABLE = False


class OutlookError(Exception):
    """Custom exception for Outlook errors."""


class OutlookTasks:
    """Helper for interacting with Outlook Tasks via COM."""

    def __init__(self, folder_path: Optional[str] = None, profile: Optional[str] = None) -> None:
        if platform.system() != "Windows" or not WIN32_AVAILABLE:
            raise OutlookError("pywin32 not available or platform not Windows")

        self.outlook = self._connect_outlook(profile)
        self.namespace = self.outlook.GetNamespace("MAPI")
        if folder_path:
            self.tasks_folder = self.namespace.GetFolderFromID(
                self._folder_id_from_path(folder_path)
            )
        else:
            self.tasks_folder = self.namespace.GetDefaultFolder(constants.olFolderTasks)

    @staticmethod
    def _connect_outlook(profile: Optional[str] = None):
        try:
            return win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except Exception:
            return win32com.client.Dispatch("Outlook.Application")

    def _folder_id_from_path(self, path: str) -> str:
        """Resolve a folder path like '\\Mailbox\\Tasks' to EntryID."""
        folder = self.namespace.Folders
        parts = [p for p in path.split("\\") if p]
        for name in parts:
            folder = folder[name]
        return folder.EntryID

    # Task operations -----------------------------------------------------

    def list_incomplete_tasks(self) -> List[Dict[str, Any]]:
        items = self.tasks_folder.Items
        items = items.Restrict("[MessageClass] = 'IPM.Task'")
        incomplete = items.Restrict("[Complete] = 0")

        results = []
        for task in incomplete:
            results.append(
                {
                    "entryId": task.EntryID,
                    "subject": task.Subject,
                    "dueDate": getattr(task, "DueDate", None),
                    "status": task.Status,
                    "body": getattr(task, "Body", None),
                }
            )
        return results

    def add_task(self, subject: str, due_date: Optional[datetime] = None, body: Optional[str] = None) -> str:
        task = self.tasks_folder.Items.Add("IPM.Task")
        task.Subject = subject
        if due_date:
            task.DueDate = due_date
        if body:
            task.Body = body
        task.Save()
        return task.EntryID

    def complete_task(self, entry_id: str) -> None:
        task = self.namespace.GetItemFromID(entry_id)
        task.Status = constants.olTaskComplete
        task.Save()

    def delete_task(self, entry_id: str) -> None:
        task = self.namespace.GetItemFromID(entry_id)
        task.Delete()


# Convenience factory ------------------------------------------------------

def get_tasks_client() -> OutlookTasks:
    folder = os.getenv("OUTLOOK_TASKS_FOLDER")
    profile = os.getenv("OUTLOOK_PROFILE")
    return OutlookTasks(folder, profile)
