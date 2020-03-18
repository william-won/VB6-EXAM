 /**********************************************************************
 *                   InterSoft International, Inc                      *
 *                        Copyright (C) 1995                           *
 ***********************************************************************
 * System:   IBM PC                                                    *
 * Program:  NETEDIT.C                                                 *
 * Author:   K.R. Robinette                                            *
 * Date:     July, 1996                                                *
 * Function: Remote Editing Support                                    *
 **********************************************************************/
#include "stdio.h"
#include "string.h"
#include "sys/types.h"
#include "sys/stat.h"

 char on[5]  = {"\033[5i"};
 char off[5] = {"\033[3i"};

 main(argc,argv)
 int argc;
 char **argv;
 {
 int  len,flag,mode;
 FILE *fd;
 char line[1024],out[1024];
 struct stat buf;
 if(argc == 2)
      {
      if((stat(argv[1],&buf)) != 0)
           {
           printf("Error, could not open %s\n",argv[1]);
           exit(-1);
           }
      mode = buf.st_mode;
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

 fd = NULL;
 system("stty -echo");
 strcpy(out,argv[1]);
 strcat(out,".new");
 line[0] = 0;
 while(1)
      {
      if(fgets(line,sizeof(line)-1,stdin) == NULL)
           {
           flag = 2;
           break;
           }
      if(line[0] == 0x02)
           {
           system("stty echo");
           printf("File was not modified\n");
           exit(0);
           }
      if(fd == NULL)
           if((fd = fopen(out,"w")) == NULL)
                {
                system("stty echo");
                printf("Error, could not open output file %s\n",out);
                exit(-3);
                }
      if(line[0] == 0x01)
           {
           flag = 1;
           break;
           }
      len = strlen(line);
      fwrite(line,1,len,fd);
      }
 if(fd)
      fclose(fd);
 if(flag == 2)
      remove(out);
      else if(flag == 1)
           {
           remove(argv[1]);
           strcpy(line,"mv ");
           strcat(line,out);
           strcat(line," ");
           strcat(line,argv[1]);
           system(line);
           chmod(argv[1],mode);
           }
 system("stty echo");
 if(flag == 1)
      printf("File was modified\n");
      else
      printf("File was not modified\n");
 }
