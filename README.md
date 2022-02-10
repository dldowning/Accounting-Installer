# xlwings installer

You can find a comprehensive documentation including video walkthrough here:  
https://docs.xlwings.org/en/latest/release.html

## Create installer

Go to `Releases` > `Draft/Create a new release`. Add a version like `1.0.0` to `Tag version`, then hit `Publish release`.

Wait a few minutes (typically between 2-10) until the installer will appear under the release. You can follow the progress under the `Actions` tab.

## Setup

### Excel file

You can add your Excel file to this repository if you like, but it is not required. Use the `xlwings release` command, to prepare the Excel file to play with the installer.

### Source code

Source code can either be embedded in the Excel file (requires an xlwings PRO license key) or added to the `src` directory here. It could, however, also be distributed via a shared folder. In this case, you'd only need to configure the `PYTHONPATH` accordingly. Embedding the code makes the distribution of updates easier as you only need to deploy the new version of the Excel file. Using the `src` directory means that you'd need to create a new version of the installer every time you have a code change.

### Data files

It is recommended to store your data files in the `data` directory. The directory gets copied next to the Python executable and can be accessed like this:

```
import sys
from pathlib import Path
DATA_DIR = Path(sys.executable).resolve().parent / 'data'
```

### Dependencies

Add your dependencies to `requirements.txt`.

If you have dependencies from **public** Git repos, you can add them by using the `https` version of the Git clone URL (`#subdirectory=<DIRECTORY>` is only required if your `setup.py` file is not in the root directory):

```
git+https://github.com/<USERNAME>/<REPO>.git@<BRANCH OR TAG OR SHA>#subdirectory=<DIRECTORY>
```

If you have dependencies from **private** Git repos, it's a bit more complicated:

**In your private repo with the source for Python package(s)**:  
* Create a deploy key via by executing the following command on a Terminal: `ssh-keygen -t rsa -b 4096 -m PEM` - it will ask you for a passphrase: Hit Enter without entering one and same for the confirmation. Note that your key needs to be in the PEM format. If you want to use an existing one, make sure it starts with `----BEGIN RSA PRIVATE KEY-----` and not with `-----BEGIN OPENSSH PRIVATE KEY-----`. You can convert an existing key into PEM format like this: `ssh-keygen -p -m PEM -f /path/to/id_rsa`.
* Go to `Settings` > `Deploy keys` and add the public key as deploy key (usually `id_rsa.pub`).


**In this repo**:  
* Use the `ssh` version of the Git clone url, however, **you need to replace the `:` with `/`**, it should look like this:

```
git+ssh://git@github.com/<USERNAME>/<REPO>.git@<BRANCH OR TAG OR SHA>#subdirectory=<DIRECTORY>
```
* Uncomment the `Install SSH key` section in `.github/workflows/main.yml`
* On this installer repo, go to `Settings` > `Secrets` and add the secret variable called `SSH_KEY`: Paste the private ssh deploy key from above (usually `id_rsa`).


### Code signing

If you sign the installer with a code sign certificate, users will see a blue `Verified Publisher` at the beginning of the installation process. Otherwise, they will see an orange `Unverified Publisher` screen. If you'd like the installer to be signed by `Zoomer Analytics GmbH`, let us know. Otherwise, you can sign the executable with your own certificate after downloading it or upload it to this repo:

* Store your code sign certificate as `sign_cert_file` in the root of this repository (make sure your repo is private).
* Go to `Settings` > `Secrets` and add the password as `code_sign_password`.

### Project details

In case you need to change the name of the installer or publisher, you can edit `.github/workflows/main.yml`:

```
PROJECT: 
APP_ID: 
APP_PUBLISHER: 
```

Comment out the `Code signing` section if you don't have a certificate.

### Python version

You can edit the Python version under `.github/workflows/main.yml`:

```
python-version: '3.8'
architecture: 'x64'
```
