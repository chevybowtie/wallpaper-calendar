When you want to replicate this environment in another system, you can run pip install, using the -r switch to specify the requirements file and pip will install the dependencies for you:

`python -m pip install -r requirements.txt`

Upgrade in the future:

`python -m pip install -U -r requirements.txt`

# Notice on Outlook integration
While this has some code to get Outlook's appointments and prepare to inject in the wallpaper, I don't use Outlook at home and I keep forgetting to set this up at work to try it. So, this feature is incomplete at best (and just outputs to console) and dangerous at worst (it *shouldn't* bork your calendar)
