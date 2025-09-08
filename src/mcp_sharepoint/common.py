import os, logging
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('mcp_sharepoint.log'), logging.StreamHandler()]
)
logger = logging.getLogger('mcp_sharepoint')

# Load environment variables
load_dotenv()

# Configuration
SHP_ID_APP = os.getenv('SHP_ID_APP')
SHP_ID_APP_SECRET = os.getenv('SHP_ID_APP_SECRET')
SHP_SITE_URL = os.getenv('SHP_SITE_URL')
SHP_DOC_LIBRARY = os.getenv('SHP_DOC_LIBRARY', 'Shared Documents/mcp_server')
SHP_TENANT_ID = os.getenv('SHP_TENANT_ID')

if not SHP_SITE_URL:
    logger.error("SHP_SITE_URL environment variable not set.")
    raise ValueError("SHP_SITE_URL environment variable not set.")
if not SHP_ID_APP:
    logger.error("SHP_ID_APP environment variable not set.")
    raise ValueError("SHP_ID_APP environment variable not set.")
if not SHP_ID_APP_SECRET:
    logger.error("SHP_ID_APP_SECRET environment variable not set.")
    raise ValueError("SHP_ID_APP_SECRET environment variable not set.")
if not SHP_TENANT_ID:
    logger.error("SHP_TENANT_ID environment variable not set.")
    raise ValueError("SHP_TENANT_ID environment variable not set.")

# Initialize MCP server
mcp = FastMCP(
    name="mcp_sharepoint",
    instructions=f"This server provides tools to interact with SharePoint documents and folders in {SHP_DOC_LIBRARY}"
)

# Initial SharePoint context
credentials = ClientCredential(SHP_ID_APP, SHP_ID_APP_SECRET)
auth_ctx = AuthenticationContext(
    f"https://login.microsoftonline.com/{SHP_TENANT_ID}"
)
sp_context = ClientContext(SHP_SITE_URL, auth_ctx).with_credentials(credentials)
