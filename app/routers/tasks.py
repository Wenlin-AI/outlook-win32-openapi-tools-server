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
        info = client.add_task(task.subject, task.dueDate, task.body)
        return info
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc


@router.post("/tasks/{task_id}/complete")
def complete_task(task_id: int):
    try:
        client = get_tasks_client()
        client.complete_task(task_id)
        return {"status": "completed"}
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc


@router.delete("/tasks/{task_id}")
def delete_task(task_id: int):
    try:
        client = get_tasks_client()
        client.delete_task(task_id)
        return {"status": "deleted"}
    except OutlookError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc
