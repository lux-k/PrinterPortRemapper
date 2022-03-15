Log file: c:\printer-remapper.txt
You have to run as admin on the system to restart services/edit registry.

Command line options:

	Option		Meaning
	------		-------
	toip		translate printer ports to IP addresses
	tohost		translate printer ports to DNS hostnames
	listports	only list the printer ports found *

	ignoresnmp	Don't worry about SNMP settings on the port *
	enablesnmp	Turn on SNMP settings for port
	disablesnmp	Turn off SNMP settings for port
			* default options

	Manually specify a newname for an old name.
	rename:oldname=newname
	
	Only change printers with IPs or hostnames that fall into this range
	range:1.2.3.4
	range:1.2.3.4-5.5.5.5	
	range:1.2.3.4,1.2.3.5

	e.g.
	PrinterPortRemapper tohost
		For any printer port
			That needs to be converted to hostname
				Do so using DNS


	PrinterPortRemapper tohost range:1.1.0.0-2.0.0.0 rename:printer1=1.2.3.4
		For any printer port
			On the network range of 1.1.0.0 through 2.0.0.0
				That needs to be converted to hostname
					Do so using dns, EXCEPT printer1 (use given name)

	To only rename one printer manually:
	PrinterPortRemapper tohost range:6.6.6.6 rename:6.6.6.6:myawesomeprinter

	To rename all printers to hostnames, given one explicit name:
	PrinterPortRemapper tohost rename:6.6.6.6:myawesomeprinter

	More than likely, if you want to impact only one printer, you want a range that matches. Otherwise you are likely
	doing more work than you thought.