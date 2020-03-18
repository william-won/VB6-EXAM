/* nted.c
 * copyright 1996 David J. Binette
 * Thu Sep 05 04:33:43 GMT 1996, dbin
 *
 * This program may ONLY be distributed with sourcecode.
 * (and any extensions you add)
 *
 * This program is FREEWARE, use it entirely at your own risk.
 * if you use it, you assume the risk and consequences of any
 * problems that may arise.
 *
 *
 * This program may be freely distributed for use with netterm.
 * Please leave my name in the credits list. (and add yours)
 * email  dbin@sce.de
 * www    http://www.sce.de
 *
 * netterm 3.0 was the current release when this program was written
 * netterm is copyright Intersoft International Inc
 * netterm email:  72060.2331@compuserve.com
 *
 *
 * This program requires Chuck Forsbergs' ZMODEM file transfer program.
 * zmodem   sz for sending files
 * zmodem   rz for receiving files
 * www  http://www.omen.com
 *
 *
 * The remote editing feature of netterm is a great idea!
 * Zmodem is a fine file transfer program.
 *
 * This program is a bit of glue to help manage remote editing
 * when accessing unix files from DOS.
 *
 * It handles the problems of DOS editors that don't understand
 * *N*X newline conventions.
 * With this program you can use "notepad" or any other plain text
 * editor to edit UNIX files locally on your PC.
 *
 * (check out James Iulianos' "lemmy" a fine VI compatible editor for DOS)
 * ( http://www.accessone.com/~jai )
 *
 *
 * This program "nted" works with netterm to manage remote file editing using
 * netterms magic <esc>[6i code sequence.
 * nted sends a file via zmodem to the remote host.
 * Then the magic sequence is sent, causing the remote host
 * to edit the received file, and lock the terminal screen.
 *
 * When the remote user is finished editing the file, the netterm program
 * sends many carriage returns and an 'rz' command
 * ( this breaks menu programs badly)
 *
 * we use Zmodems -a option to handle DOS cr/nl mapping when we send the file.
 * we wait for the remote side to return the file, then we
 * convert it back to unix format
 * This program also handles pathnames properly allowing you to edit
 * files that are not in the current directory.
 *
 * the program returns 1 if there is any error
 *                     0 if all went well
 *
 *
 * TAB STOPS ON THIS FILE are set to 4 spaces
 *
 * If your *NIX does not have strrchr(), try rindex()
 */

/*
 * Revision history
 * Tue Sep 10 08:53:16 GMT 1996, dbin
 * added command line options -a -b -v and -z"xxx"
 * added multiple file editing
 * handled case where the received filename
 *         is converted to all upper or lower case
 */

#include <stdio.h>
#include <string.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <ctype.h>
#include <stdlib.h>
#include <unistd.h>

/* =================================================== */

#define CNV_NONE 0
#define CNV_2NIX 1
#define CNV_2DOS 2

/* default Zmodem option is -y overwrite destination file */
#define ZMODEM_OPTS     "-y"

/* =================================================== */

char *Progname;

/* =================================================== */

/* display usage information and exit */
void usage(void)
{
        fprintf(stderr,"usage: %s [options] filename\n", Progname);
        fprintf(stderr,"See nted.c for program documentation and usage precautions.\n");
        fprintf(stderr,"NOT TO BE DISTRIBUTED WITHOUT SOURCECODE\n");
        fprintf(stderr,"Program author: David J. Binette September 10 1996\n");
        fprintf(stderr,"-a      transfer file as ascii\n");
        fprintf(stderr,"-b      transfer file as binary\n");
        fprintf(stderr,"-v      verbose output\n");
        fprintf(stderr,"-z\"xxx\" zmodem options\n");
        exit(1);
}

/* =================================================== */

/* copy an ASCII file
 * line length limit 8Kb
 * if the dos2nix flag is set: cr/nl is converted to nl on output
 * if the dos2nix flag is clr: no special processing id done
 *
 * return 1 if error
 *        0 if all ok
 */

int copyfile(char *source, char *dest, int convert)
{
        FILE *ifp;
        FILE *ofp;
        char buf[8194];
        int n;

        if((ifp = fopen(source,"r")) == (FILE*)0)
        {
                perror(source);
                return 1;
        }
        if((ofp = fopen(dest,"w")) == (FILE*)0)
        {
                perror(dest);
                fclose(ifp);
                return 1;
        }

        while(fgets(buf, 8192, ifp))
        {
                if(convert == CNV_2NIX)         /* strip CR (dos to nix conversion) */
                {
                        if((n = strlen(buf)) > 1)
                        {
                                if((buf[n-2] == '\r') && (buf[n-1] == '\n'))
                                {
                                        buf[n-2] = '\n';
                                        buf[n-1] = '\0';
                                }
                        }
                }
                else
                if(convert == CNV_2DOS)         /* add CR (nix to dos conversion) */
                {
                        if((n = strlen(buf)) > 1)
                        {
                                if(buf[n-1] == '\n')
                                {
                                        buf[n-1] = '\r';
                                        buf[n]   = '\n';
                                        buf[n+1] = '\0';
                                }
                        }
                }

                if(fputs(buf,ofp) == EOF)
                {
                        perror(dest);
                        fclose(ofp);
                        fclose(ifp);
                        return 1;
                }
        }
        fclose(ofp);
        fclose(ifp);

        return 0;
}

/* =================================================== */
/* return base name of path/file */
char * basename(char *s)
{
        char *p = s;
        if((p = strrchr(s,'/')) != (char*)0)
                ++p;
        else
                p=s;
        return p;
}

/* =================================================== */

void main(int argc, char **argv)
{
        char * origpathfile;            /* the original complete path/filename */
        char * filename;                        /* the filename without any leading path */
        char * origdir;                         /* current working directory */
        char * zopt = ZMODEM_OPTS;      /* default is to overwrite files */
        char newpathfile[1024];         /* the new complete path/filename */
        char tmpdir[1024];                      /* a temporary working dir */
        char buf[1024];                         /* a command line buffer */
        int     verbose = 0;                    /* verbose mode */
        int c;
        int retval = 0;
        int convert = 1;                        /* default 1 = send as ascii files */
        extern char *optarg;            /* for getopt */
        extern int optind;                      /* for getopt */

        Progname = basename(argv[0]);

        while((c = getopt(argc, argv, "abvz:")) != -1)
        {
                switch (c) {
                case 'a': convert = 1;          break;
                case 'b': convert = 0;;         break;
                case 'v': verbose = 1;;         break;
                case 'z': zopt = optarg;        break;
                case '?': usage();                      break;
                }
        }

        if(optind == argc)
                usage();

        /* determine the current directory name */
        if((origdir = getcwd(NULL, 2048)) == NULL)
        {
                perror("pwd");
                exit(1);
        }

        /* make a working directory */
        sprintf(tmpdir,"/tmp/nted.%d", getpid());
        if(mkdir(tmpdir, S_IEXEC | S_IREAD | S_IWRITE))
        {
                perror(tmpdir);
                exit(1);
        }

        /* now process all of the specified files */
        /* each file is sent, edited and retrieved seperately */

        for( ; optind < argc; optind++)
        {

                /* split the original filename */
                origpathfile    = argv[optind];
                filename                = basename(argv[optind]);

                /* copy the file to the temporary directory */
                sprintf(newpathfile,"%s/%s", tmpdir, filename);

                if(copyfile(origpathfile, newpathfile, convert ? CNV_2DOS : 0))
                {
                        unlink(newpathfile);
                        rmdir(tmpdir);
                        exit(1);
                }

                /* make the temp directory the current directory */
                if(chdir(tmpdir))
                {
                        perror(tmpdir);
                        unlink(newpathfile);
                        rmdir(tmpdir);
                        exit(1);
                }

                /* Use Zmodem to transfer the file */
                if(verbose)
                        fprintf(stderr,"Sending file to PC for editing\n");
                sprintf(buf, "sz %s %s\n",zopt, filename);
                if((retval = system(buf)) != 0)
                        fprintf(stderr,"%s: warning! system returned %d\n", retval);

                /* trash the copy, we dont need it anymore */
                unlink(newpathfile);

                /* send netterm magic */
                /* causes netterm to invoke editor on received file */
                fputs("\033[6i", stderr);
                fflush(stderr);

                /* now wait for some kind of user input
                 * or the other side to send back the file
                 */

                retval = 1;
                while(1)
                {
                        if(verbose)
                                fprintf(stderr,
                                        "You are remotely editing \"%s\"\n",
                                        origpathfile);
                        fprintf(stderr,"type OK to continue\n");
                        fflush(stderr);

                        buf[0] = '\0';
                        if(fgets(buf,sizeof(buf), stdin) == NULL)
                                break;

                        buf[0] = (char)tolower(buf[0]);
                        buf[1] = (char)tolower(buf[1]);

                        if( (buf[0] == 'o') && (buf[1] == 'k') )
                                break;

                        /* receive the edited file */
                        if((buf[0] == 'r') && (buf[1] == 'z'))
                        {
                                if(verbose)
                                        fprintf(stderr,
                                                        "Now receiving updated file from remote editor\n");
                                system("rz");
                                retval = 0;
                                break;
                        }
                }

                /* go back to the original directory */
                chdir(origdir);

                if(retval == 0) /* copy the file to its destination */
                {
                        int filerror = 0;                       /* filename error? on receive */

                        if(access(newpathfile,F_OK))
                        {
                                /* the expected file was NOT found */
                                /* convert it to lowercase and try again */
                                char *p;
                                p = basename(newpathfile);
                                while(p && *p)
                                {
                                        *p = (char)tolower(*p);
                                        p++;
                                }
                        }
                        if(access(newpathfile,F_OK))
                        {
                                /* the expected file was NOT found */
                                /* convert it to uppercase and try again */
                                char *p;
                                p = basename(newpathfile);
                                while(p && *p)
                                {
                                        *p = (char)toupper(*p);
                                        p++;
                                }
                        }
                        if(access(newpathfile,F_OK))
                        {
                                filerror = 1;
                                fprintf(stderr,
                                                "Sorry, the remote PC returned an unknown file\n");
                                fprintf(stderr,
                                                "look in %s\n",tmpdir);
                        }
                        else
                        {
                                filerror = copyfile(newpathfile,
                                                        origpathfile,
                                                        convert ? CNV_2NIX : 0) != 0;
                        }
                        if(filerror)
                        {
                                fprintf(stderr,"press [ENTER] to continue");
                                fgets(buf,sizeof(buf), stdin);
                        }
                }

                unlink(newpathfile);
        }
        rmdir(tmpdir);
        exit(0);
}

/* =================================================== */
