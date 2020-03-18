 /**********************************************************************
 *                   InterSoft International, Inc                      *
 *                        Copyright (C) 1995                           *
 ***********************************************************************
 * System:   IBM PC                                                    *
 * Program:  NETSEND.C                                                 *
 * Author:   K.R. Robinette                                            *
 * Date:     January, 1996                                             *
 * Function: Remote Printing Support                                   *
 *           Supports HP Printer Escapes and long lines.               *
 **********************************************************************/
#include "stdio.h"
#include "string.h"

 char on[5]  = {"\033[5i"};
 char off[5] = {"\033[4i"};

 main(argc,argv)
 int argc;
 char **argv;
 {
 int  len,flag;
 FILE *fd;
 char line[1024];
 if(argc == 2)
      {
      if((fd = fopen(argv[1],"r")) == NULL)
           {
           printf("Error, could not open %s\n",argv[1]);
           exit(-1);
           }
      if((fwrite(on,1,4,stdout)) != 4)
           {
           printf("Error, writing to network\n");
           exit(-1);
           }
      while(1)
           {
           if(fgets(line,1023,fd) == NULL)
                break;
           len = strlen(line);
           fwrite(line,1,len,stdout);
           }
      fwrite(off,1,4,stdout);
      }
      else
           {
           printf("Input filename required\n");
           exit(-2);
           }
 fclose(fd);
 }
