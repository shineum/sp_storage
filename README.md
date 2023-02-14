# Installation

Use pip:

```
pip install git+https://github.com/shineum/sp_storage.git
```


# Getting Started
Set the following in your settings.py:

```
SHAREPOINT_TENANT = '<tenant name>'
SHAREPOINT_TENANT_ID = '<tenant id - uuid format>'
SHAREPOINT_CLIENT_ID = '<client id - uuid format>'
SHAREPOINT_CLIENT_SECRET = '<client secret>'
SHAREPOINT_USERNAME = '<username>'
SHAREPOINT_PASSWORD = '<password>'
SHAREPOINT_SITE_NAME = '<site name>'
SHAREPOINT_ROOT_DIR = '<root dir>'
SHAREPOINT_BLOB_MAX_MEMORY_SIZE = '<blob file size - defulat 16M>'
```

> note: 
> If 'SHAREPOINT_CLIENT_ID' and 'SHAREPOINT_CLIENT_SECRET' are provided 'SHAREPOINT_USERNAME' and 'SHAREPOINT_PASSWORD' will be ignored.


# Usage

## Fields
Set the default storage as 'sp_storage.sharepoint.SharepointStorage'
or import SharepointStorage and set storage for each models.

```
DEFAULT_FILE_STORAGE = 'sp_storage.sharepoint.SharepointStorage'

# model
from django.db import models

class Sample(models.Model):
    pdf = models.FileField(upload_to='pdfs')

```

or

```
# model
from sp_storage.sharepoint import SharepointStorage
from django.db import models

class Sample(models.Model):
    pdf = models.FileField(upload_to='pdfs', storage=SharepointStorage)
```

## Storage
```
>>> from sp_storage.sharepoint import SharepointStorage
>>> storage = SharepointStorage()

>>> storage.exists('storage_test')
False

>>> file = storage.open('storage_test', 'w')
>>> file.write('storage test')
>>> file.close()

>>> storage.exists('storage_test')
True

>>> file = storage.open('storage_test', 'r')
>>> file.read()
b'storage test'
>>> file.close()

>>> storage.delete('storage_test')
>>> storage.exists('storage_test')
False
```