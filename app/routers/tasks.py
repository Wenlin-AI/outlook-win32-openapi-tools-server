"""CRUD routes for Outlook Tasks."""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from ..outlook import OutlookError, get_tasks_client

router = APIRouter()


class TaskCreate(BaseModel):
    subject: str
    dueDate: Optional[datetime] = None
    body: Optional[str] = None


@router.get("/tasks")
def list_tasks():
    try:
        client = get_tasks_client()
        return client.list_incomplete_tasks()
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc


@router.post("/tasks")
def create_task(task: TaskCreate):
    try:
        client = get_tasks_client()
        entry_id = client.add_task(task.subject, task.dueDate, task.body)
        return {"entryId": entry_id}
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc


@router.post("/tasks/{entry_id}/complete")
def complete_task(entry_id: str):
    try:
        client = get_tasks_client()
        client.complete_task(entry_id)
        return {"status": "completed"}
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc


@router.delete("/tasks/{entry_id}")
def delete_task(entry_id: str):
    try:
        client = get_tasks_client()
        client.delete_task(entry_id)
        return {"status": "deleted"}
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc
