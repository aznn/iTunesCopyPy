import sys, win32com.client, os, subprocess

def readCommands(argv):
    # Implement command line options
    pass

def displayMenu(itunes):
    print '\nScript to copy an iTunes playlist to a speicified directory.'
    print 'Select a playlist to move :'

    args = {}
    playlists = itunes.LibrarySource.Playlists

    i = 0
    for playlist in playlists:
        i += 1
        print '%d - %s' % (i, playlist.Name)

    print '\n> ',
    choice = int(raw_input()) - 1
    args['playlist'] = playlists[choice]

    print 'Playlist %s selected!' % playlists[choice].Name

    print '\nEnter Directory to move playlist to. Eg: C:\music\ (INCLUDE the trailing \ !)'
    print 'Dir : ',
    dir = raw_input()
    args['dir'] = dir

    return args


def movePlaylist(itunes, playlist, dir):
    print 'Moving playlist(%s) to Dir(%s), enter 1 to confirm : ' % (playlist.Name, dir),
    if raw_input() != '1' : sys.exit(0)

    # Verify dir
    if not os.path.exists(dir):
        print 'Directory dosent exist, create folder? (0/1) : ',
        if raw_input() == '0': sys.exit(0)

        os.makedirs(dir)

    failed = []
    for track in playlist.Tracks:
        trackPath = track.Location
        filename = os.path.basename(trackPath)
        copyPath = dir
        
        copy = [
                'copy',
                trackPath,
                copyPath
            ]

        devnull = open('/dev/null', 'w')
        subprocess.call(copy, shell=True, stdout=devnull)

        if os.path.isfile(copyPath + filename):
            print 'Success : %s' % filename
        else:
            print 'Failed : %s' % filename
            failed.append(track)

    fail = len(failed)
    if fail: print '\n%d Failed!' % fail
    else: print '\nAll Files successfully copied!'


def runScript():
    itunes = win32com.client.Dispatch("iTunes.Application")
    args = displayMenu(itunes)

    movePlaylist(itunes, **args)

"""
if __name__ == "__main__":
    runScript()

"""

runScript()