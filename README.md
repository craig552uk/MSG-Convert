# MSG Convert

A perl script to convert Outlook .msg files in to .eml files to read on Ubuntu Linux using Thunderbird.

Originally by Matijs van Zuijlen [matijs@matijs.net](matijs@matijs.net)
Later edits by Craig Russell [craig@craig-russell.co.uk](craig@craig-russell.co.uk)

## Usage Instructions

1. Install the required perl packages with the following commands    

    sudo perl -MCPAN -e 'install("Email::Outlook::Message")'
    sudo perl -MCPAN -e 'install("Email::LocalDelivery")'
    sudo perl -MCPAN -e 'install("Getopt::Long")'
    sudo perl -MCPAN -e 'install("Pod::Usage")'
    sudo perl -MCPAN -e 'install("File::Basename")'
    
2. Install Thunderbird (if you don;t already have it) with `sudo apt-get install thunderbird`

3. Download and save `msgconvert.pl` to your computer

4. Make it executable with `chmod +x /path/to/msgconvert.pl`

5. Open .msg files with `/apth/to/msgconvert.pl /path/to/emailfile.msg`    

You can configure Nautilus to open `.msg` files when you double-click them    

1. Right click any .msg file

2. Select *Open with Other Application*

3. Click the arrow next to *Use a Custom Command*

4. Click the *Browse* button and choose the `msgconvert.pl` file

Now when you double click a `.msg` file, it will be automatically converted to a `.eml` file and opened in Thunderbird.
