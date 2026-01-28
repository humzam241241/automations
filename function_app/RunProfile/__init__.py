"""
Azure Function: HTTP trigger to run a profile.
POST /api/run?profile=<name>
"""
import json
import logging
import sys
from pathlib import Path

import azure.functions as func

# Add parent to path for imports
sys.path.append(str(Path(__file__).resolve().parents[2]))

from core.engine import ExecutionEngine
from core.profile_loader import ProfileLoader
from function_app.shared.graph_client import GraphClient
from function_app.shared.config_loader import load_app_settings


def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Run a profile via HTTP trigger.
    
    Query params:
        - profile: Profile name (required)
        - explain: If 'true', include rule match explanations
    """
    logging.info("RunProfile endpoint invoked.")
    
    profile_name = req.params.get("profile")
    explain = req.params.get("explain", "false").lower() == "true"
    
    if not profile_name:
        return func.HttpResponse(
            json.dumps({"error": "Missing 'profile' query parameter"}),
            status_code=400,
            mimetype="application/json",
        )
    
    # Load settings
    try:
        settings = load_app_settings()
    except Exception as e:
        logging.exception("Failed to load settings")
        return func.HttpResponse(
            json.dumps({"error": f"Configuration error: {e}"}),
            status_code=500,
            mimetype="application/json",
        )
    
    # Get Graph token
    try:
        graph = GraphClient(
            tenant_id=settings["tenant_id"],
            client_id=settings["client_id"],
            client_secret=settings["client_secret"],
        )
        token = graph.get_access_token()
    except Exception as e:
        logging.exception("Failed to acquire Graph token")
        return func.HttpResponse(
            json.dumps({"error": f"Authentication failed: {e}"}),
            status_code=500,
            mimetype="application/json",
        )
    
    # Load profile
    loader = ProfileLoader()
    profile = loader.load_profile(profile_name)
    
    if not profile:
        return func.HttpResponse(
            json.dumps({"error": f"Profile '{profile_name}' not found"}),
            status_code=404,
            mimetype="application/json",
        )
    
    # Validate profile
    errors = loader.validate_profile(profile)
    if errors:
        return func.HttpResponse(
            json.dumps({"error": "Invalid profile", "validation_errors": errors}),
            status_code=400,
            mimetype="application/json",
        )
    
    # Execute
    try:
        engine = ExecutionEngine(access_token=token)
        result = engine.run_profile(profile, explain=explain)
        
        logging.info(f"Profile '{profile_name}' executed successfully: {result}")
        
        return func.HttpResponse(
            json.dumps(result),
            status_code=200,
            mimetype="application/json",
        )
    except Exception as e:
        logging.exception(f"Failed to execute profile '{profile_name}'")
        return func.HttpResponse(
            json.dumps({"error": f"Execution failed: {e}"}),
            status_code=500,
            mimetype="application/json",
        )
