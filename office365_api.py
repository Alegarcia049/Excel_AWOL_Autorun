import environ
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from ofice365.sharepoint.files.file import file

env = environ.Env()
environ.Env().read_env()

USSERNAME = env("sharepoint_email")
PASSWORD = env("sharepoint_password")
SHAREPOINT_SITE = env("sharepoint_url_site")
SHAREPOINT_SITE_NAME = env("sharepoint_site_name")
SHAREPOINT_DOC = env("sharepoint_doc_library")

class SharePoint:
    def _auth(self);
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
            USSERNAME,
            PASSWORD
            )
        )
        return conn
    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by