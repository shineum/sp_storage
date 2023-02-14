import os
import mimetypes
from tempfile import SpooledTemporaryFile

from django.core.exceptions import ImproperlyConfigured
from django.core.files.base import File
from django.utils.deconstruct import deconstructible

from storages.base import BaseStorage
from storages.utils import clean_name
from storages.utils import setting
from storages.utils import to_bytes

try:
    from office365.runtime.auth.authentication_context import AuthenticationContext
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.webs.web import Web
    from office365.sharepoint.files.file import File as sharepoint_file

except ImportError:
    raise ImproperlyConfigured("Could not load Sharepoint Storage bindings.\n"
                               "See https://pypi.org/project/Office365-REST-Python-Client")

@deconstructible
class SharepointStorageFile(File):

    def __init__(self, name, mode, storage):
        self.name = name
        self._mode = mode
        self._storage = storage
        self._file = None
        self._is_dirty = False
        self.mime_type = mimetypes.guess_type(name)[0]

    @property
    def size(self):
        return self._file.size

    def _get_file(self):
        if self._file is not None:
            return self._file
            
        file = SpooledTemporaryFile(
            max_size=self._storage.max_memory_size,
            suffix=".SharepointStorageFile",
            dir=setting("FILE_UPLOAD_TEMP_DIR", None))

        if self._storage.exists(self.name):
            sp_file_path = self._storage.get_sp_file_path(self.name)
            sp_file = self._storage.service_context.web.get_file_by_server_relative_url(sp_file_path)
            sp_file.download(file).execute_query()
        if 'r' in self._mode:
            file.seek(0)

        self._file = file
        return self._file

    def _set_file(self, value):
        self._file = value

    file = property(_get_file, _set_file)

    def read(self, *args, **kwargs):
        if 'r' not in self._mode:
            raise AttributeError("File was not opened in read mode.")
        return super().read(*args, **kwargs)

    def write(self, content):
        if 'w' not in self._mode:
            raise AttributeError("File was not opened in write mode.")
        self._is_dirty = True
        return super().write(to_bytes(content))

    def close(self):
        if self._file is None:
            return
        if self._is_dirty:
            self._file.seek(0)
            self._storage._save(self.name, self._file)
            self._is_dirty = False
        self._file.close()
        self._file = None


@deconstructible
class SharepointStorage(BaseStorage):

    def __init__(self, **settings):
        super().__init__(**settings)
        self._service_context = None

    def get_default_settings(self):
        tenant = setting("SHAREPOINT_TENANT")
        site_name = setting("SHAREPOINT_SITE_NAME")
        sharepoint_url = f"https://{tenant}.sharepoint.com"
        site_url = f"{sharepoint_url}/sites/{site_name}"        
        use_app_auth = setting("SHAREPOINT_CLIENT_ID") and setting("SHAREPOINT_CLIENT_SECRET")

        return {
            "tenant": tenant,
            "tenant_id": setting("SHAREPOINT_TENANT_ID"),
            "client_id": setting("SHAREPOINT_CLIENT_ID"),
            "client_secret": setting("SHAREPOINT_CLIENT_SECRET"),
            "username": setting("SHAREPOINT_USERNAME"),
            "password": setting("SHAREPOINT_PASSWORD"),
            "sharepoint_url": sharepoint_url,
            "site_name": site_name,
            "site_url": site_url,
            "use_app_auth": use_app_auth,
            "root_dir": setting("SHAREPOINT_ROOT_DIR"),
            "max_memory_size": setting("SHAREPOINT_BLOB_MAX_MEMORY_SIZE", 16*1024*1024),
        }

    def __get_sharepoint_context_using_app(self):
        client_credentials = ClientCredential(self.client_id, self.client_secret)
        return ClientContext(self.site_url).with_credentials(client_credentials)

    def __get_sharepoint_context_using_user(self):
        user_credentials = UserCredential(self.username, self.password)
        return ClientContext(self.site_url).with_credentials(user_credentials)

    def __get_service_context(self):
        if self.use_app_auth:
            return self.__get_sharepoint_context_using_app()
        else:
            return self.__get_sharepoint_context_using_user()

    @property
    def service_context(self):
        if self._service_context is None:
            self._service_context = self.__get_service_context()
        return self._service_context

    def _open(self, name, mode="rb"):
        return SharepointStorageFile(name, mode, self)

    def get_relative_url(self, name):
        return "/".join([self.root_dir, name]) if self.root_dir else name

    def get_relative_dir(self, name):
        rurl = self.get_relative_url(name)
        sp_file_dir_chunks = rurl.split("/")
        return "/".join(sp_file_dir_chunks[:-1])
        
    def get_sp_file_path(self, name):
        return '/'.join(['/sites', self.site_name, self.get_relative_url(name)])

    def get_raw_resource_uri(self, name):
        return '/'.join([self.site_url, self.get_relative_url(name)])

    def exists(self, name):
        try:
            ctx = self.service_context
            sp_file_path = self.get_sp_file_path(name)            
            sp_file = ctx.web.get_file_by_server_relative_url(sp_file_path).get().execute_query()
            return sp_file.exists
        except Exception as e:
            print(e)
        return False

    def delete(self, name):
        try:
            ctx = self.service_context
            sp_file_path = self.get_sp_file_path(name)
            sp_file = ctx.web.get_file_by_server_relative_url(sp_file_path)
            sp_file.recycle().execute_query()
        except Exception as e:
            print(e)

    def size(self, name):
        ret = -1
        try:
            ctx = self.service_context
            sp_file_path = self.get_sp_file_path(name)
            sp_file = ctx.web.get_file_by_server_relative_url(sp_file_path).get().execute_query()
            return sp_file.length
        except Exception as e:
            print(e)
        return ret

    def create_dir(self, sp_file_dir):
        sp_file_dir_chunks = sp_file_dir.split("/")
        if len(sp_file_dir_chunks) > 1:
            self.create_dir("/".join(sp_file_dir_chunks[:-1]))
        try:
            ctx = self.service_context
            ctx.web.folders.add(sp_file_dir).execute_query()
        except Exception as e:
            print (e)

    def _save(self, name, content):
        cleaned_name = clean_name(name)
        try:
            ctx = self.service_context
            sp_file_dir = self.get_relative_dir(name)
            file_name = os.path.basename(name)
            if self.exists(name):
                self.delete(name)
            self.create_dir(sp_file_dir)
            target_folder = ctx.web.get_folder_by_server_relative_url(sp_file_dir)
            target_file = target_folder.upload_file(file_name, content).execute_query()
            target_file.checkin('', 1).execute_query()
        except Exception as e:
            print(e)
        return cleaned_name

    def url(self, name):
        try:
            ctx = self.service_context
            link_url = self.get_raw_resource_uri(name)
            client_result = Web.create_organization_sharing_link(ctx, link_url, False).execute_query()
            return client_result.value
        except Exception as e:
            print(e)
        return name