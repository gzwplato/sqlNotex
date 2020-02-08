# sqlNotex

sqlNotex is a multi-platform software useful to manage a large amount of textual notes in Markdown format on a single computer or in a local network using the open source database Firebird [firebird.sql](firebirdsql.org).
The notes are divided into notebooks and, within them, in sections, and it is possible to associate to each note a list of activities to be done, a series of attachments (files of any kind), tags and links to other notes. The search function may find the wished notes starting from the title, the text content, the modification date, the tags, the name of the attachments or activities.

The text of the notes can be formatted if it is written by the user in Markdown format. The titles are displayed in bold and with a larger font than the rest of the text, while the various markers (asterisk, slash, etc.) are formatted with their own color. It is also possible to copy the text of a note with any possible activities in HTML format and paste it into a word processor, or display it in the browser, or automatically insert it into a new LibreOffice Writer document, thus obtaining a regularly formatted document.

sqlNotex has been written with Lazarus (www.lazarus-ide.org) and a modified version of the RichMemo component [wiki.freepascal.org/RichMemo](wiki.freepascal.org/RichMemo), whose modified source code is included in the source code of sqlNotex, and accesses the Firebird database through the Zeos components [sourceforge.net/projects/zeoslib](sourceforge.net/projects/zeoslib).
sqlNotex is free software, as it is released under the GPL version 3 license or following, available on www.gnu.org/licenses/gpl-3.0.html, which the user must accept in order to use it.

The official website of the software is https://sites.google.com/view/sqlnotex/.

See the [Wiki](https://github.com/maxnd/sqlNotex/wiki) for the user manual of the software.

Before contributing to the development of this software, see [the Contributing file](https://github.com/maxnd/sqlNotex/blob/master/CONTRIBUTING.md).
