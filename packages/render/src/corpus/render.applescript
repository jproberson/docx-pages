-- Has Word export its own pdf of each document handed to it, one line of answer
-- per document, and never stops on one that will not open.
--
-- This is the same request `authored/render.applescript` makes, asked of other
-- people's documents rather than of ours, and they are a harder crowd:
--
-- **`display alerts` has to be off.** A real document opens with something to say
-- often enough that leaving them on is not an option: Word puts up a dialog, the
-- whole run dies with `-1712` however long the timeout, and the document is left
-- open with a `~$` lock file beside it that poisons every later run. Off, the same
-- documents export without a word.
--
-- **Every document is timed and caught on its own.** One that will not open is a
-- fact worth reporting rather than the end of the batch, and a batch is worth
-- keeping short so that a Word wedged past saving costs one batch and not the run.
--
-- Word is deliberately left running at the end: a cold start costs far more than
-- every export in a batch put together. The caller quits it when there is no batch
-- left to run.

on run argv
	set out to ""
	tell application "Microsoft Word"
		set display alerts to none
		repeat with each in argv
			set startedAt to current date
			with timeout of 600 seconds
				try
					set d to open file name POSIX file each with read only
					set target to (text 1 thru -6 of each) & ".pdf"
					save as d file format format PDF file name POSIX file target
					try
						close d saving no
					end try
					set out to out & "ok|" & ((current date) - startedAt) & "|" & each & linefeed
				on error message number code
					try
						close d saving no
					end try
					set out to out & "fail|" & code & "|" & each & linefeed
				end try
			end timeout
		end repeat
	end tell
	return out
end run
