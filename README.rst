Timestat
=============
A cron-job to statistics your time expendirture.

Versions
--------
Timestat works with Python 2.7

Usage
-----

::

    $ Timestat.py [or Make Timestat as a cron-jor in crontab]


Details
--------
Please add the Timestat into /usr/bin  and change it mode

mv Timestat.py Timestat

sudo chmod 755 Timestat

sudo chown root Timestat

sudo chgrp root Timestat

sudo mv Timestat /usr/bin

Then add one command into your /etc/crontab file(as bellow):

#m h   dom mon dow  user	command 

0 20	*	* 	*	root	/usr/bin/Timestat  

		Jan 10,2019. 
			Chengdu, China.
