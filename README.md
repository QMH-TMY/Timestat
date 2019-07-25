------------
# Timestat #
------------
A cron-job to statistics your time expendirture.

# Requirement #
Python 3

# Usage #
    $ python3 Timestat.py [or Make Timestat as a cron-jor in crontab]


# Note #
Please add the Timestat into /usr/bin  and change it mode

    $ mv Timestat.py Timestat
    $ sudo chmod 755 Timestat
    $ sudo chown root Timestat
    $ sudo chgrp root Timestat
    $ sudo mv Timestat /usr/bin
    Then add one command into your /etc/crontab file(as bellow):
    #m h   dom mon dow  user	command 
    0 20	*	* 	*	root	/usr/bin/Timestat  
