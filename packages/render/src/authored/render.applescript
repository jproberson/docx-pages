-- Has Word render each authored document to a pdf beside it, so that a rendering
-- of the same document can be put next to Word's own and looked at.
--
-- Word will not export into a sandboxed directory without a dialog that blocks the
-- whole script, but it exports into the project's own without complaint.

on run argv
	set out to ""
	tell application "Microsoft Word"
		repeat with each in argv
			set d to open file name POSIX file each
			set target to (text 1 thru -6 of each) & ".pdf"
			save as d file format format PDF file name POSIX file target
			set out to out & target & linefeed
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
