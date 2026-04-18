"""
Party Master Service for CMA / DPR Builder.
Handles logic for saving, loading, and searching party-wise projects.
"""

import os
import json
import logging
import uuid
from datetime import datetime
from typing import List, Optional

from services.cma.models import CmaProject, PartyProfile

logger = logging.getLogger(__name__)

class PartyMasterService:
    """Service to manage CMA project persistence."""
    
    @staticmethod
    def get_storage_path() -> str:
        """Returns the directory where CMA projects are stored."""
        path = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 
                            'BKL_Office_Tools', 'cma_projects')
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    @classmethod
    def save_project(cls, project: CmaProject) -> str:
        """
        Saves or updates a CMA project to a JSON file.
        Returns the file path.
        """
        if not project.party_id:
            project.party_id = str(uuid.uuid4())[:8].upper()
            project.created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        project.updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Use business name and party_id for filename for uniqueness and readability
        safe_name = "".join(x for x in project.profile.business_name if x.isalnum() or x in " -_").strip()
        if not safe_name:
            safe_name = "Untitled_Party"
        
        filename = f"{safe_name}_{project.party_id}.json"
        file_path = os.path.join(cls.get_storage_path(), filename)
        
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(project.to_dict(), f, indent=2, ensure_ascii=False)
            logger.info(f"Project saved: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"Failed to save project: {e}")
            raise RuntimeError(f"Could not save project: {e}")

    @classmethod
    def load_project(cls, file_path: str) -> CmaProject:
        """Loads a CMA project from a JSON file."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Project file not found: {file_path}")
            
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return CmaProject.from_dict(data)
        except Exception as e:
            logger.error(f"Failed to load project: {e}")
            raise RuntimeError(f"Could not load project: {e}")

    @classmethod
    def list_projects(cls) -> List[dict]:
        """
        Returns a list of summary dictionaries for all saved projects.
        Useful for dashboard listings.
        """
        projects = []
        storage_path = cls.get_storage_path()
        
        for filename in os.listdir(storage_path):
            if filename.endswith(".json"):
                full_path = os.path.join(storage_path, filename)
                try:
                    # We only read the basic info for the list to keep it fast
                    with open(full_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        profile = data.get("profile", {})
                        projects.append({
                            "party_id": data.get("party_id"),
                            "business_name": profile.get("business_name", "Untitled"),
                            "pan": profile.get("pan", ""),
                            "entity_type": profile.get("entity_type", ""),
                            "category": profile.get("business_category", ""),
                            "updated_at": data.get("updated_at", ""),
                            "file_path": full_path
                        })
                except Exception as e:
                    logger.warning(f"Skipping corrupt project file {filename}: {e}")
                    
        # Sort by last updated
        projects.sort(key=lambda x: x["updated_at"], reverse=True)
        return projects

    @classmethod
    def delete_project(cls, party_id: str) -> bool:
        """Deletes a project by matching its party_id in the filename."""
        storage_path = cls.get_storage_path()
        for filename in os.listdir(storage_path):
            if party_id in filename and filename.endswith(".json"):
                os.remove(os.path.join(storage_path, filename))
                return True
        return False
