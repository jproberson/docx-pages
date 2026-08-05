-- Asks Word where everything in each authored document landed, and closes each one
-- behind it. Word answers layout questions far better than reading its pdf back:
-- it reports empty paragraphs, which draw nothing and so cannot be measured from a
-- rendering at all, and it says which page each one landed on.
--
-- Every document is opened and closed inside the one run. A document left open
-- makes every later `open file name` fail, which reads like a broken script and is
-- not, so the close is tried and the whole run quits Word behind itself.

on run argv
	set out to ""
	tell application "Microsoft Word"
		repeat with each in argv
			set d to open file name POSIX file each with read only
			set out to out & "DOC|" & each & linefeed

			repeat with i from 1 to (count of paragraphs of d)
				set r to text object of paragraph i of d
				set p to get range information r information type active end page number
				set v to get range information r information type vertical position relative to page
				set h to get range information r information type horizontal position relative to text boundary
				set out to out & "P|" & i & "|" & p & "|" & v & "|" & h & linefeed
			end repeat

			-- A shape reports the size Word settled on, which for a box that fits
			-- itself to its text is the whole question being asked.
			repeat with j from 1 to (count of shapes of d)
				set s to shape j of d
				set out to out & "S|" & (name of s) & "|" & (width of s) & "|" & (height of s) & linefeed
			end repeat

			try
				close d saving no
			end try
		end repeat
	end tell

	try
		tell application "Microsoft Word" to quit saving no
	end try
	return out
end run
