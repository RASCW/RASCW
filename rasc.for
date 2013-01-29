      PROGRAM RASC
C
C           FOR RANKING AND SCALING OF BIOSTRATIGRAPHIC EVENTS
C                         (VERSION 15 - 1996)
C
C
C            RRR        A       SSS      CCC          PPPP      CCC
C            R   R     A A     S   S    C   C         P   P    C   C
C            R   R    A   A    S        C             P   P    C
C            RRRR     AAAAA     SSS     C      =====  PPPP     C
C            R R      A   A        S    C             P        C
C            R  R     A   A    S   S    C   C         P        C   C
C            R   R    A   A     SSS      CCC          P         CCC
C
C
C
C                                   by
C
C             F.P.Agterberg, F.M.Gradstein, L.D.Nel, S.N.Lew,
C             M.Heller, W.S.Gradstein, M.D'Iorio, D.Gillis & Z.Huang
C
C             The present Ansi F77 version 14 was compiled with the
C             Lahey F77L FORTRAN Language System, version 5.01. It has
C             output direction, which allows the user to specify in the
C             parameters (input) file which of 11 results modules will
C             be saved in the "main output" file and which in the "extra
C             output" file. This feature is useful for quick perusal of
C             the calculated zonation (Optimum Sequence and its Scaling).
C             In RASC version 14 (=RASC14) the maximum size of data files
C             is: 50 wells, 998 dictionary taxa, 150 taxa per well,
C             20 unique (rare) events and 20 marker horizons.
C             The (scaled) optimum sequence can include up to 145 taxa,
C             depending on the value of the IOCR input parameter.
C             Details on the parameters file and data files format for
C             RASC14 are given below under Program Inputs.
C
C             RASC-PC was originally developed using IBM Professional
C             Fortran, version 1.00. The switch to Lahey F77L, 5.01 was
C             made by F.P. Agterberg in June 1995. It involves a change
C             in the .INP file of which the first three records used to
C             be in free format. Now this entire file is fixed format.
C
C
C  REFERENCES:
C  -----------
C              F.P.Agterberg & L.D.Nel, 1982a. Algorithms for the
C              ranking of stratigraphic events. Computers and Geosciences,
C              vol.8,no.1,p.69-90.
C
C              F.P.Agterberg & L.D.Nel, 1982b. Algorithms for the
C              scaling of stratigraphic events. Computers and Geosciences,
C              vol.8,no.1, p.163-189.
C
C              F.M.Gradstein & F.P.Agterberg, 1982. Models of Cenozoic
C              foraminiferal stratigraphy, Northwestern Atlantic Margin. In:
C              Cubitt, J.M. and Reyment, R.A. (eds), Quantitative
C              Stratigraphic Correlation. J.Wiley & Sons, U.K., p. 119-173.
C
C              F.M.Gradstein, F.P.Agterberg, J.C.Brower & W.Schwarzacher,
C              1985. Quantitative Stratigraphy. Reidel Publ.Co. & Unesco,
C              598p.
C
C              F.P.Agterberg, 1990. Automated Stratigraphic Correlation.
C              Elsevier, Amsterdam, 424p.
C
C              W.S.Gradstein, 1993. RASC12A MANUAL. Available with this
C              program.
C
C
C    PROGRAM INPUTS
C    --------------
C
C     RECORD NO. 1     RUN PARAMETERS     (FIXED FORMAT)
C     ------------     --------------
C
C   NS     -  (INTEGER)   NUMBER OR SEQUENCES OR WELLS  (NS <= 50)
C   IOCR   -  (INTEGER)   ELEMENTS OCCURRING FEWER THAN "IOCR" TIMES IN
C                             THE DATA SET WILL BE IGNORED
C   INIQ   -  (INTEGER)   = 1, IF UNIQUE EVENTS OR MARKER HORIZONS ARE
C                                   INCLUDED  (RECORDS 4 AND 5)
C   ITER   -  (INTEGER)   MAXIMUM NUMBER OF CUMULATIVE ORDER MATRIX
C                             TRANSFORMATIONS ALLOWED  (EXAMPLE: 12000)
C   CRIT1  -  (REAL)      TRANSPOSE ELEMENTS WITH SUM LESS THAN "CRIT1"
C                             IN THE ORDER MATRIX WILL BE ZEROED BEFORE
C                             THE RANKING SOLUTION.  (CRIT1 <= IOCR)
C   TOL    -  (REAL)      TOLERANCE:  S(I,J) MAY BE LESS THAN S(J,I)
C                                     BY AS MUCH AS "TOL"
C   AAA    -  (REAL)      FRACTILE FOR TRUNCATION POINT OF NORMAL
C                             DISTRIBUTION.  (EXAMPLE: AAA = 1.645)
C   CRIT2  -  (REAL)      TRANSPOSE ELEMENTS WITH SUM LESS THAN "CRIT2"
C                             IN THE ORDER MATRIX WILL BE ZEROED BEFORE
C                             THE SCALING ANALYSIS.  (CRIT2 >= CRIT1)
C
C
C     RECORD NO. 2     PROCESSING CONTROL       (FIXED FORMAT)
C     ------------     ------------------
C
C  ALL 14 PARAMETERS ON THIS RECORD ARE INTEGERS.
C  FOR SHORT VERSION OF RASC (RANKING ALGORITHMS ONLY), IALPHA=0
C
C
C   ITAPE     = 1   FOR DATA TO BE READ FROM "TAPE10"
C              ELSE,       DATA WILL BE READ FROM RECORDS IMMEDIATELY
C                          FOLLOWING RECORDS NO. 4 OR 5
C   IOMAT     = 1   FOR PRINTOUT OF ORDER AND FREQUENCY MATRICES
C                     AS WELL AS INTERMEDIATE TABLES;
C              ELSE,       THESE OUTPUTS WILL BE SUPPRESSED
C   ISRT      = 0,  MODIFIED HAY METHOD WITHOUT PRESORTING
C               1,  DATA WILL BE PRE-SEQUENCED FOR OPTIMIZED STARTING
C                     SEQUENCE
C              ELSE,       CONDENSED OPTIMUM SEQUENCE AFTER PRESORTING
C   IALPHA    = 0,  TERMINATION AFTER RANKING SOLUTION;
C             = 1,  SCALING ANALYSIS WILL BE DONE;
C              ELSE,       TERMINATION AFTER RANKING SOLUTION, BUT
C                          STEPWISE SEQUENCING PROGRESS WILL BE
C                          PRINTED BEFORE TERMINATION.
C   ITAB1     = 1,  AN OCCURRENCE TABLE FOR THE WELLS IS TO BE PRINTED
C              ELSE,       NO TABLE.
C   ISCORE    = 1,  STEP MODEL COMPARISON OF INDIVIDUAL WELLS AND
C                       FOSSILS WITH OPTIMUM SEQUENCE IS PERFORMED
C   ICOMP     = 1   FOR NORMALITY TESTS ON INDIVIDUAL WELLS
C   ISKIP     = 1   IF CUMULATIVE ORDER MATRIX IS TO BE USED
C                    (RANKING SOLUTION WILL BE BASED ON PRESORTING
C                     ONLY)
C              ELSE,       RASC WILL GO AHEAD AND PERFORM MATRIX
C                          PERMUTATIONS.
C   IFIN      = 1   FOR APPLICATION OF FINAL RE-ORDERING
C   INOSC     = 0,  NO SCALING OUTPUT;
C             = 1,  WEIGHTED SCALING OUTPUT ONLY;
C             = 2,  WEIGHTED AND UNWEIGHTED SCALING OUTPUT.
C   INEG      = 1,  LARGE DISTANCES FOR SMALL SAMPLES WILL BE SUPPRESSED
C              ELSE,       NO SUPPRESSION.
C   ISCAT     = 1,  SCATTERGRAMS;
C             = ELSE,      NO SCATTERGRAMS
C   IVAR      = 1,   VARIANCE ANALYSIS TO BE PERFORMED FOR EACH WELL;
C              ELSE,       NO VARIANCE ANALYSIS
C   ICASC     = 0, NO WELL DATA OUTPUT FILE;
C             = 1, OUTPUT FILE FOR USE AS INPUT TO ERRORBAR PROGRAM;
C              ELSE,       WELL DATA OUTPUT FILE.
C
C     RECORD NO. 3     OUTPUT DIRECTION (FIXED FORMAT)
C     ------------     ----------------
C
C     THIS RECORD IS MADE UP OF 11 INTEGER VALUES (FREE FORMAT)
C     WHICH DIRECT OUTPUT FROM THE 11 DIFFERENT PROGRAM SECTIONS.
C     THESE VALUES ARE STORED IN THE PROGRAM IN THE ARRAY OUT().
C     THE FIRST VALUE REFERS TO SECTION 1, THE SECOND TO SECTION
C     2 AND SO ON. IF THE VALUE IS SET TO 1, THE CORRESPONDING SECTION'S
C     OUTPUT WILL BE DIRECTED TO THE MAIN OUTPUT FILE. OTHERWISE
C     THE SECTIONS OUTPUT WILL GO TO THE EXTRA OUTPUT FILE.
C
C     OUT(1) -  TABULATION OF EVENT OCCURRENCES (DISABLED IN RASC14)
C     OUT(2) -  MODIFIED SEQUENCE DATA (DISABLED IN RASC14)
C     OUT(3) -  DICTIONARIES
C     OUT(4) -  CYCLES
C     OUT(5) -  OPTIMUM SEQUENCE
C     OUT(6) -  OCCURRENCE TABLE AND STEP MODEL
C     OUT(7) -  SCATTERGRAMS AND VARIANCE ANALYSIS
C     OUT(8) -  SCALING (WEIGHTED)
C     OUT(9) -  SCALING AFTER 5 ITERATIONS
C     OUT(10)-  NORMALITY TEST
C     OUT(11)-  UNIQUE (RARE) EVENTS IN OPTIMUM SEQUENCE
C
C
C     RECORD NO. 4     UNIQUE (RARE) EVENTS
C     ------------     -------------
C
C     IF INIQ = 1, UP TO 20 UNIQUE EVENTS WILL BE READ FROM THIS
C     RECORD IN 20I4 FORMAT.  IF NO UNIQUE EVENTS ARE REQUESTED,
C     THIS RECORD IS LEFT BLANK.  IF UNIQUE EVENTS ARE INCLUDED,
C     "ICOMP" MUST EQUAL 1.
C
C
C     RECORD NO. 5     MARKER HORIZONS
C     ------------     ---------------
C
C     IF INIQ = 1, UP TO 20 MARKER HORIZONS WILL BE READ FROM
C     THIS RECORD IN 20I4 FORMAT.  IF NO MARKER HORIZONS ARE
C     REQUESTED, THIS RECORD IS LEFT BLANK.
C
C
C
C     DATA SET              (FIXED FORMAT)
C     --------
C
C     OBSERVED EVENTS:      THE OBSERVED DATA ARE GIVEN AS "NS"
C                           SEQUENCES.  A "SEQUENCE" CONSISTS OF
C            A TITLE RECORD (IN 5A4 FORMAT) FOLLOWED BY ONE OR MORE
C            RECORDS OF SEQUENCE DATA (IN MULTIPLE I4 FORMAT).
C                EACH OF THESE DATA RECORDS IS PARTITIONED INTO 20
C            FIELDS, EACH FIELD HAVING I4 FORMAT.  THUS, THERE
C            MAY BE UP TO 20 EVENTS IN ONE RECORD.
C                NOTE:   FIELDS NEED NOT BE FILLED CONSECUTIVELY.
C                        EMPTY FIELDS ARE IGNORED.
C
C                A SEQUENCE IS DEEMED TO HAVE ENDED WHEN A FIELD
C            CONTAINING   -999   IS ENCOUNTERED.
C                THE SEQUENCES ARE READ EITHER FROM "TAPE10"
C            (IF ITAPE = 1) OR FROM RECORDS IMMEDIATELY FOLLOWING
C            RECORD NUMBER 4  (IF ITAPE = 0).
C
C
C     DICTIONARY:           THE DICTIONARY IS MERELY A LIST OF
C                           EVENT LABELS:  ONE LABEL PER RECORD,
C            10A4 FORMAT.  AN EXTRA RECORD WITH THE LABEL 'LAST'
C            IN THE FIRST 4 COLUMNS MUST BE PLACED AT THE END.
C            LIMIT:  998 RECORDS  (NOT COUNTING 'LAST' RECORD)
C            THE DICTIONARY IS ALWAYS READ FROM "TAPE99".
C
C  ---------------------------------------------------------------------
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (kdim7=25, kdim8=30, MAXCYC=300, MAXUQ=50)
C
      COMMON N, NS, MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON  /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      COMMON  /DELTA/ IOCR, INIQ, CRIT1, TOL, AAA, CRIT2, MAX, ITER,
     +        IOMAT, ISRT, IALPHA, ITAB1, ISCORE, ICOMP, ISKIP, IFIN,
     +        INOSC, INEG, ISCAT, ivar, icasc, TMAT(KDIM1,KDIM3),ift
      COMMON  /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER AA, AID, COUNT, T, TEST, IPOS(KDIM6), IQD(KDIM3),
     +        IRANGE(KDIM6,2), JIRCOD(KDIM6), IPAIR(2,KDIM6+MAXUQ),
     +        MPAIR(KDIM6+MAXUQ), TMAT, OUT(11), NRPT,ivec(kdim4),
     +        iqcyc(kdim5),icode2(kdim5),icode3(kdim5),ircodea(kdim4),
     +        iqcyc2(kdim7,kdim5),ircode2(kdim4),iqdar(kdim6+maxuq),
     +        iqcyc3(kdim7,kdim8),iqcyc4(kdim7,kdim8),ircodeo(kdim4),
     +        iutem(maxuq,2),jvan(kdim6,4),iunc(kdim4,2),icon(kdim4),
     +        nopt(kdim4)
      REAL    A(300), B(300), CC(600), QDAR(KDIM6+MAXUQ), WVEC(KDIM5),
     +        XLEV(KDIM6+MAXUQ),rmat(kdim3,3),runiq(kdim4,2),
     +        dev(kdim6+maxuq),bindev(kdim6),sd(kdim1),af(kdim1),
     +        bf(kdim1),cf(kdim1),sdopt(kdim4),st(kdim6+maxuq)
      character     istar2*3
      character*2   logo,res2,res3
      CHARACTER*4   NAME, ITITLE,res1
      CHARACTER*40  FOSSIL(KDIM4)
      CHARACTER*50  INPFIL,DATFIL,DICFIL,deqfil,batfil
      character*45 OUTFIL
      character*50 OUTFILa,outfilb,outfilc
      character*50 outfild, outfile,outfilf,outfilg,outfilp,outfilq
      character*50 grafcum,grafden,grafsc1,grafsc2,grafva1,grafra1
      character*50 grafva2,grafra2,summary,optseq
      character*50 tabfl1,tabfl2,tabfl3,tabfl0,tabuni,tabwel
      character*50 tabocc,tabpen,tabnor,tabcyc,taball
      character*10 result
      character*8 result4
      character*80 contents
      logical nodon
      istar2 = '**'
C
C  EXPLANATION OF CONSTANTS IN THE PARAMETER LIST
C
C  KDIM1  - MAXIMUM NUMBER OF WELLS (SECTIONS)
C  KDIM3  - MAXIMUM LENGTH OF ANY ONE WELL
C  KDIM4  - MAXIMUM NUMBER OF EVENT LABELS IN THE DICTIONARY
C  KDIM5  - MAXIMUM NUMBER OF FOSSILS PERMITTED TO BE
C             IMPLICATED IN ONE CYCLE
C  KDIM6  - MAXIMUM NUMBER OF EVENTS REMAINING AFTER FILTERING
c  kdim7  - maximum number of cycles in condensed optimum sequence (COS)
c  kdim8  - maximum number of cycling events per cluster in COS
C  MAXCYC - MAXIMUM NUMBER OF CYCLES ALLOWED
C  MAXUQ  - MAXIMUM NUMBER OF UNIQUE (RARE) EVENTS
C
C  DISPLAY PROGRAM NAME

      WRITE(*,'(4(/))')
      WRITE(*,'(12X,A)')
     &'RRRR       A       SSS      CCC          PPPP      CCC'
      WRITE(*,'(12X,A)')
     &'R   R     A A     S   S    C   C         P   P    C   C'
      WRITE(*,'(12X,A)')
     &'R   R    A   A    S        C             P   P    C'
      WRITE(*,'(12X,A)')
     &'RRRR     AAAAA     SSS     C      =====  PPPP     C'
      WRITE(*,'(12X,A)')
     &'R R      A   A        S    C             P        C'
      WRITE(*,'(12X,A)')
     &'R  R     A   A    S   S    C   C         P        C   C'
      WRITE(*,'(12X,A)')
     &'R   R    A   A     SSS      CCC          P         CCC'
      WRITE (*,*)' '
      write(*,*) ' '
      write(*,*) ' '
      WRITE (*,'(12X,A)')
     &'         RASC VERSION 18 (2002)'
      write (*,*) ' '
      write(*,'(12x,a)')
     &'                    by'
      write (*,*) ' '
      write(*,'(12x,a)')
     &'         F.P. Agterberg and F.M. Gradstein'
      WRITE (*,*)' '
      write (*,*)' '
c      WRITE (*,'(5X,A)')
c     &'Enter name for input file with run parameters (e.g. *.inp): '
c      READ (*,900) INPFIL
  900 FORMAT (A50)
c      WRITE (*,'(5X,A)')
c     &'Enter name for input file with well sequence (e.g. *.dat): '
c      READ (*,900) DATFIL
c      WRITE (*,'(5X,A)')
c     &'Enter name for input file with dictionary taxa (e.g. *.dic): '
c      READ (*,900) DICFIL
c      WRITE (*,*) '  '
c      WRITE (*,'(5X,A)')
c     &'Enter name for main output file (e.g. *m.out): '
c     &'Enter name (1-7 characters long) for output files (e.g. 20sep): '
c      READ (*,900) OUTFIL1
c      read (*,9011) outfil
 9011 format(a7)
c      WRITE (*,'(5X,A)')
c     &'Enter name for extra output file (e.g. *e.out): '
c      READ (*,900) OUTFIL2
c      WRITE (*,'(5X,A)')
c     &'Enter name of supplementary plot data file (e.g. *p.out): '
c      READ (*,900) T7FILE
c      write (*,'(5x,a)')
c     &'Enter name of variance analysis file (e.g. *.var): '
c      read (*,900) tape24
c      write (*,'(5x,a)')
c     &'Enter name of EXCEL/CASC input file (e.g. *.xlw): '
c      read (*,900) TAPE25
c      TAPE25='excel.inp'
C
      open(3,file='rasctemp',status='unknown')
c      open(4,file='rascflag',status='unknown')
      read(3,1998) inpfil
      read(3,1998) datfil
      batfil=datfil
      read(3,1998) dicfil
      read(3,1997) outfil
 1998 format(a50)
 1997 format(a45)
      OPEN (5,FILE=INPFIL,BLANK='ZERO')
      OPEN (10,FILE=DATFIL,STATUS='OLD',BLANK='ZERO')
      OPEN (99,FILE=DICFIL,STATUS='OLD',ACCESS='SEQUENTIAL')
c      open (22, file='tape22',status='unknown')
c      write(22,902) outfil,outfil,outfil,outfil,outfil,outfil,outfil,
c     +outfil,outfil
c  902 format(a5,'a.out',a5,'b.out',a5,'c.out',a5,'d.out',a5,'e.out',a5,
c     +'f.out',a5,'g.out',a5,'p.out',a5,'q.out')
c      rewind 22
c      read(22,903) outfila,outfilb,outfilc,outfild,outfile,outfilf,
c     +outfilg,outfilp,outfilq
c  903 format(9a10)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
c      if(outfil(7:7).eq.' ') outfil=' '//outfil(1:6)
      do 10101 i=1,50
      if(outfil(45:45).eq.' ') outfil=' '//outfil(1:44)
10101 continue
      outfila=outfil//'a.out'
c      write(*,*) outfil,outfila
      outfilb=outfil//'b.out'
      outfilc=outfil//'c.out'
      outfild=outfil//'d.out'
      outfile=outfil//'e.out'
      outfilf=outfil//'f.out'
      outfilg=outfil//'g.out'
      outfilp=outfil//'h.out'
      outfilq=outfil//'i.out'
      grafcum=outfil//'.cum'
      grafden=outfil//'.den'
      grafsc1=outfil//'.sc1'
      grafsc2=outfil//'.sc2'
      deqfil=outfil//'.dep'
c      seqfil=outfil//'x.dat'
      grafva1=outfil//'.va1'
      grafva2=outfil//'.va2'
      grafra1=outfil//'.ra1'
      grafra2=outfil//'.ra2'
      summary=outfil//'.sum'
      optseq=outfil//'.opt'
      tabfl1=outfil//'.fl1'
      tabfl2=outfil//'.fl2'
      tabfl3=outfil//'.fl3'
      tabfl0=outfil//'.fl0'
      tabuni=outfil//'.uni'
      tabwel=outfil//'.wel'
      tabocc=outfil//'.occ'
      tabpen=outfil//'.pen'
      tabnor=outfil//'.nor'
      tabcyc=outfil//'.cyc'
      taball=outfil//'.all'
      do 10102 i=1,50
      if(outfila(1:1).eq.' ') outfila=outfila(2:50)//' '
      if(outfilb(1:1).eq.' ') outfilb=outfilb(2:50)//' '  
      if(outfilc(1:1).eq.' ') outfilc=outfilc(2:50)//' '
      if(outfild(1:1).eq.' ') outfild=outfild(2:50)//' '  
      if(outfile(1:1).eq.' ') outfile=outfile(2:50)//' '
      if(outfilf(1:1).eq.' ') outfilf=outfilf(2:50)//' '  
      if(outfilg(1:1).eq.' ') outfilg=outfilg(2:50)//' '
      if(outfilp(1:1).eq.' ') outfilp=outfilp(2:50)//' '  
      if(outfilq(1:1).eq.' ') outfilq=outfilq(2:50)//' '
      if(grafcum(1:1).eq.' ') grafcum=grafcum(2:50)//' '  
      if(grafden(1:1).eq.' ') grafden=grafden(2:50)//' '
      if(grafsc1(1:1).eq.' ') grafsc1=grafsc1(2:50)//' '  
      if(grafsc2(1:1).eq.' ') grafsc2=grafsc2(2:50)//' '
      if(deqfil(1:1).eq.' ') deqfil=deqfil(2:50)//' '  
c      if(seqfil(1:1).eq.' ') seqfil=seqfil(2:50)//' '
      if(grafva1(1:1).eq.' ') grafva1=grafva1(2:50)//' '
      if(grafva2(1:1).eq.' ') grafva2=grafva2(2:50)//' '  
      if(grafra1(1:1).eq.' ') grafra1=grafra1(2:50)//' '
      if(grafra2(1:1).eq.' ') grafra2=grafra2(2:50)//' '  
      if(summary(1:1).eq.' ') summary=summary(2:50)//' '
      if(optseq(1:1).eq.' ') optseq=optseq(2:50)//' '
      if(tabfl1(1:1).eq.' ') tabfl1=tabfl1(2:50)//' '
      if(tabfl2(1:1).eq.' ') tabfl2=tabfl2(2:50)//' '
      if(tabfl3(1:1).eq.' ') tabfl3=tabfl3(2:50)//' '
      if(tabfl0(1:1).eq.' ') tabfl0=tabfl0(2:50)//' '
      if(tabuni(1:1).eq.' ') tabuni=tabuni(2:50)//' '
      if(tabwel(1:1).eq.' ') tabwel=tabwel(2:50)//' '
      if(tabocc(1:1).eq.' ') tabocc=tabocc(2:50)//' '
      if(tabpen(1:1).eq.' ') tabpen=tabpen(2:50)//' '
      if(tabnor(1:1).eq.' ') tabnor=tabnor(2:50)//' '
      if(tabcyc(1:1).eq.' ') tabcyc=tabcyc(2:50)//' '
      if(taball(1:1).eq.' ') taball=taball(2:50)//' '
10102 continue
      OPEN (61,FILE=OUTFILa,STATUS='UNKNOWN')
      OPEN (62,FILE=OUTFILb,STATUS='UNKNOWN')
      OPEN (7,FILE=outfilp,ACCESS='SEQUENTIAL',STATUS='UNKNOWN')
      open(23,file='tape23',form='unformatted')
      open(25,file=outfilq,status='unknown',access='sequential')
      open(63,file=outfild,status='unknown')
      open(66,file=outfilf,status='unknown')
      open(64,file=outfilc,status='unknown')
      open(65,file=outfile,status='unknown')
      open(67,file=outfilg,status='unknown')
      open(70,file=grafcum,status='unknown')
      open(71,file=grafden,status='unknown')
      open(72,file=grafsc1,status='unknown')
      open(73,file=grafva1,status='unknown')
      open(76,file=grafva2,status='unknown')
      open(74,file=grafra1,status='unknown')
      open(77,file=grafra2,status='unknown')
      open(75,file=grafsc2,status='unknown')
      open(80,file=summary,status='unknown')
      open(81,file=optseq,status='unknown')
      open(82,file=deqfil,status='unknown')
      open(83,file='seqfil',status='unknown')
      open(84,file=tabfl1,status='unknown')
      open(85,file=tabfl2,status='unknown')
      open(86,file=tabfl3,status='unknown')
      open(87,file=tabfl0,status='unknown')
      open(88,file=tabuni,status='unknown')
      open(89,file=tabwel,status='unknown')
      open(90,file=tabocc,status='unknown')
      open(91,file=tabpen,status='unknown')
      open(92,file=tabnor,status='unknown')
      open(93,file=tabcyc,status='unknown')
      open(94,file=taball,status='unknown')
      call date_and_time(result4)
      res1=result4(1:4)
      res2=result4(5:6)
      res3=result4(7:8)
      result=res1//'/'//res2//'/'//res3
      write(61,905) outfila,result
      write(62,905) outfilb,result
      write(64,905) outfilc,result
      write(63,905) outfild,result
      write(65,905) outfile,result
      write(66,905) outfilf,result
      write(67,905) outfilg,result
      write(7,905) outfilp,result
      write(7,906)
c
comment added on 11/11/2003
c
      maxcycc=2*maxcyc
c
  906 format(/' EXTRA OUTPUT FILE.  POSSIBLE PLOTTING PROGRAM INPUT'/)
C
C  READS IN:  INPUT (RUN PARAMETERS), TAPE10 (SEQUENCE DATA),
C             TAPE99 (DICTIONARY)
C  CREATES:   OUTPUT (THE PRINT FILE OF RESULTS)
C             TAPE7 (INPUT TO 'DENO' PROGRAMME FOR DISSPLA PLOTS OF
C                    OPTIMUM FOSSIL SEQUENCE AND DENDROGRAMS)
C
C
C
C
C
C     WRITE A FRONT PAGE ONTO THE OUTPUT FILE
      CALL FPAGE
C
C     READ INPUTS
C     -----------
C
      itony=0
      CALL READIN (inew,iutem,OUT)
      WRITE (*,'(5X,A)') 'Thank you.   Please wait '//
     & '(or Ctrl-Break to abort)'
c      if(inosc.ne.0) inosc=2
      if (icasc.eq.1) then
      write(25,905) outfilq,result
      write(25,907)
      endif
      if(icasc.eq.0) write(25,9088)
 9088 format(' FILE IS EMPTY BECAUSE ICASC = 0')
  907 format(/' EXTRA OUTPUT FILE.  REQUIRED AS INPUT FILE FOR CASC'/)
  905 format (16x,a50,2x,a10)
C THE LINES BETWEEN C... AND C... WERE ADDED BY Z. HUANG IN JULY, 1992
C...
      NODIC=0
  100 NODIC = NODIC + 1
      NWELL = 0
      DO 170 I = 1,NS
          DO 150 J = 1, KDIM3
            IF (NODIC.EQ.ABS(TMAT(I,J))) THEN
            N1 = N1 + 1
            IF (NWELL.EQ.0) THEN
            N2 = N2 + 1
            NWELL = 1
            ENDIF
            ENDIF
  150     CONTINUE
  170    CONTINUE
      IF (NODIC.NE.N) GO TO 100
C...
C
C  CONSTRUCT  NUNIQ(KDIM1)  - MAP SHOWING WHICH SEQUENCES HAVE
C                               UNIQUE EVENTS
C  IRCODE(KDIM4)  - BEING USED HERE AS A TABLE TO BE SENT TO
C                     OCCTAB().  (REALLY ONLY NEED KDIM6 OF THE
C                     KDIM4 AVAILABLE.)
C

      DO 205 I = 1,NS
          DO 200 J = 1,KDIM3
              ID = TMAT(I,J)
              AID = IABS (ID)
              IF (AID.EQ.0) GO TO 205
              IF (INIQ.EQ.1 .AND. IUNIQ(AID,1).EQ.1) NUNIQ(I) = 1
              IRCODE(AID) = IRCODE(AID) + 1
  200     CONTINUE
  205 CONTINUE
      if(iniq.eq.1) then
      if(iutem(1,1).ne.0) then
      write(out(1),1060)
      write(89,1060)
 1060 format(/////' RECORDS FOR THE FOLLOWING UNIQUE (RARE) EVENTS ',
     +'HAVE BEEN SELECTED:'/)
      write(out(1),1061)
      write(89,1061)
      endif
 1061 format(' NUMBER   FREQUENCY       NAME'/)
      do 105 i=1,20
      id=iutem(i,1)
      if(id.eq.0) goto 110
      write(out(1),1070) id,ircode(id),(ititle(id,j),j=1,10)
      write(89,1070) id,ircode(id),(ititle(id,j),j=1,10) 
  105 continue
 1070 format(3x,i3,7x,i2,10x,10a4)
  110 if(iutem(1,2).ne.0) then
      write(out(1),1080)
 1080 format(///' THE FOLLOWING MARKER HORIZONS HAVE BEEN SELECTED:'/)
      write(out(1),1061)
      endif
      do 115 i=1,20
      id=iutem(i,2)
      if(id.eq.0) goto 120
      write(out(1),1070) id,ircode(id),(ititle(id,j),j=1,10)
  115 continue
      endif
  120 continue
C
C  PRINT A TABLE SHOWING FOR EACH DICTIONARY EVENT THE NUMBER
C  OF WELLS IN WHICH THIS EVENT APPEARS
C
      CALL OCCTAB (OUT(1),iocr)
C
C
C  LOW OCCURRENCE FILTERING AND RECODING OF DATA
C
C  ALL EVENTS IN THE DATA SET WHICH DO NOT OCCUR IN AT LEAST
C  "IOCR" SEQUENCES ARE ELIMINATED, AND THE DATA ARE RE-CODED SUCH
C  THAT IF "MMAX" EVENTS ARE RETAINED, THE NEW CODE NUMBERS WILL
C  RUN FROM 1 TO "MMAX".
C  "TMAT" WILL BECOME THE FILTERED (BUT NOT RE-CODED) DATA SET.
C
      CALL HPFILT (TMAT,IOCR,INIQ,IOMAT,OUT(2))
      MOPTI=MMAX
c      if(ialpha.eq.1.and.iutem(1,1).eq.0) write(71,10070) mopti,mopti
10070 format(/' NUMBERS OF EVENTS:',3i4)
C
C
C  PRESORT OPTION:  A PRELIMINARY SEQUENCE IS DERIVED FROM THE SORT-
C  ING OF EVENT 'SCORES' BASED ON THE FREQUENCIES OF ALL EVENTS,
C  COMPARED WITH ALL OTHER EVENTS IN AN ORDER RELATION MATRIX.
C
      IF (ISRT.NE.0) CALL PRESRT (IOMAT,OUT(2))

C  PRINT THE DICTIONARY TO OUTPUT FILE

      NCOLS = 10
      CALL ASORT (NCOLS,IPOS,FOSSIL)
      write(out(3),7038)
      write(out(3),7039)
 7038 format('DICTIONARY OF EVENTS')
 7039 format('____________________')
      WRITE (out(3),7040)
 7040 FORMAT (////10X,'NUMERICAL LISTING',30X,'ALPHABETIC LISTING'/)
      DO 250 I = 1,N
          WRITE (out(3),7050) I, (ITITLE(I,J), J = 1,10),IPOS(I),
     +                        FOSSIL(I)
 7050     FORMAT (1X, I5, 2X, 10A4, 1X, I5, 2X, A40)
  250 CONTINUE



C
C
C
C     RANKING SOLUTION
C     ----------------
C
C
C  CREATION OF CUMULATIVE ORDER MATRIX
C
C  CUMULATIVE ORDER MATRIX C(I,J) IS CONSTRUCTED SUCH THAT
C  ELEMENTS C(I,J) CONTAIN THE NUMBER OF TIMES EVENT I
C  OCCURRED ABOVE (BEFORE) EVENT J.
C
      WRITE (out(4),3000)
      WRITE (out(4),3001)
 3000 FORMAT (////'EVENT CYCLES PRIOR TO RANKING')
 3001 FORMAT ('_____________________________')
      DO 305 I = 1,KDIM6
          DO 300 J = 1,KDIM6
              C(I,J) = 0.0
  300     CONTINUE
  305 CONTINUE
      WRITE (out(4),3010) IOCR, INT (CRIT1)
 3010 FORMAT (//' RUN FOR', I2, ' OR MORE OCCURRENCES AND', I2,
     + ' OR MORE PAIRS')
C
C  PRODUCES A NEW C.O. MATRIX BASED ON THE CURRENT STATE OF THE
C  WELL DATA.  (THAT IS,  PRESORTED OR ...)
C
C  IX  - THE RECODED DATASET
C
      DO 325 L = 1,NS
          J = 1
  307     IF (J.GT.KDIM3)  GO TO 325
          MM = IX(L,J)
          IF (MM.EQ.0)  GO TO 325
              TEST = 0
              K = J
              I = IABS (MM)
  310         K = K + 1
              IF (K.GT.KDIM3)  GO TO 320
              AA = IX(L,K)
              IF (AA.EQ.0) GO TO 320
                  KK = IABS (AA)
                  IF (AA.LT.0 .AND. TEST.LE.0) GO TO 315
                      TEST = TEST + 1
                      C(I,KK) = C(I,KK) + 1.0
                      GO TO 310
  315             C(I,KK) = C(I,KK) + 0.5
                      C(KK,I) = C(KK,I) + 0.5
                      GO TO 310
  320         J = J + 1
              GO TO 307
  325 CONTINUE
      IF (IOMAT.NE.1) GO TO 331
          WRITE (out(4),3020)
 3020     FORMAT (/////,24H CUMULATIVE ORDER MATRIX/)
          DO 330 I = 1,MMAX
              WRITE (out(4),3030)
 3030         FORMAT (//)
              WRITE (out(4),3040) (C(I,J), J = 1,MMAX)
  330     CONTINUE
 3040     FORMAT (1X, 20F6.1)
  331 ia=1
c      mfmax=mmax
      do 459 i=1,kdim7
      do 458 j=1,kdim5
      iqcyc2(i,j)=0
  458 continue
  459 continue
      goto 335
  461 do 462 i=1,mmax
        do 463 j=1,icyc
        if(iqcyc(j).eq.icode(i)) then
        icode2(j)=i
        icode3(j)=i
        endif
  463   continue
  462 continue
c The following two write statements were helpful for debugging
c     condensed optimum sequence option
c      write(*,*) (icode2(j),j=1,icyc)
      ib=ia/2
      do 464 j=1,icyc
      iqcyc2(ib,j)=iqcyc(j)
  464 continue
c      write(*,*) ib,(iqcyc2(ib,j),j=1,icyc)
      kk1=icode2(1)
      do 446 k=2,icyc
      kk=icode2(k)
      do 447 i=1,mmax
      do 448 j=1,mmax
      if(j.eq.kk) c(i,kk1)=c(i,kk1)+c(i,kk)
  448 continue
  447 continue
  446 continue
      do 466 k=2,icyc
      kk=icode2(k)
      do 467 j=1,mmax
      do 468 i=1,mmax
      if(i.eq.kk) c(kk1,j)=c(kk1,j)+c(kk,j)
  468 continue
  467 continue
  466 continue
      c(kk1,kk1)=0.0
      mmaxx=mmax
      do 476 k=2,icyc
      kk=icode2(k)
      k1=k+1
      if(icode2(k1).ge.kk) icode2(k1)=icode2(k1)-1
      do 477 i=1,mmax
      mmax2=mmax-1
      do 478 ii=1,mmax2
      if(ii.ge.kk) then
      ii1=ii+1
      c(i,ii)=c(i,ii1)
      if(i.eq.1) icode(ii)=icode(ii1)
      endif
  478 continue
  477 continue
      mmax=mmax2
  476 continue
      do 481 i=1,icyc
      icode2(i)=icode3(i)
  481 continue
      mmax=mmaxx
      do 486 k=2,icyc
      kk=icode2(k)
      k1=k+1
      if(icode2(k1).ge.kk) icode2(k1)=icode2(k1)-1
      do 487 j=1,mmax
      mmax2=mmax-1
      do 488 ii=1,mmax2
      if(ii.ge.kk) then
      ii1=ii+1
      c(ii,j)=c(ii1,j)
      endif
  488 continue
  487 continue
      mmax=mmax2
  486 continue
      goto 400
C
C  MODIFICATION OF CUMULATIVE ORDER MATRIX
C
C  THE TRANSPOSE ELEMENT PAIRS C(I,J) AND C(J,I)
C  WHOSE SUM IS LESS THAN CRIT1 ARE ZEROED
C
  335 IKNT = 0
      IMAX = MMAX - 1
      DO 345 I = 1,IMAX
          L = I + 1
          DO 340 J = L,MMAX
              IF ((C(I,J)+C(J,I)) .GE. CRIT1) GO TO 340
              C(I,J) = 0.0
              C(J,I) = 0.0
              IKNT = IKNT + 1
  340     CONTINUE
  345 CONTINUE
      MMSQ = (MMAX * MMAX - MMAX) * 0.5
      WRITE (out(4),3050)
 3050 FORMAT (//' MODIFICATION OF ORDER MATRIX:')
      WRITE (out(4),3060) CRIT1, IKNT, MMSQ
 3060 FORMAT (/' BASED ON CRIT1 =',F5.1,',',I5,' PAIRS OUT OF',I6,
     + ' HAVE BEEN ZEROED')
      IF (IOMAT.NE.1) GO TO 400
          WRITE (out(4),3070)
 3070     FORMAT (/////35X,'MODIFIED RELATION MATRIX'/)
          DO 350 I = 1,MMAX
              WRITE (out(4),3030)
              WRITE (out(4),3080) (C(I,J), J = 1,MMAX)
  350     CONTINUE
 3080     FORMAT (/20F6.1)
C
C  OPTIMUM SEQUENCE DETERMINED BY MATRIX TRANSFORMATION (RANKING)
C
C  AN OPTIMUM SEQUENCE IS DETERMINED BY EXAMINING THE ORDER MATRIX.
C  FREQUENCIES (TRANSPOSE ELEMENTS C(I,J) AND C(J,I) ) ARE
C  COMPARED AND ROWS AND COLUMNS I AND J ARE INTERCHANGED SUCH THAT
C  ALL LARGER ELEMENTS APPEAR IN THE UPPER TRIANGLE OF THE MATRIX
C
C  COUNT  - # OF ITERATIONS
C  ICORT  - (A SWITCH;  SENT TO CYCLE SUBRTN.)
C  ICYC   - # FOSSILS IN THE CYCLE;  SENT TO CYCLE SUBRTN.
C
  400 DO 405 I = 1,KDIM6
          IPOS(I) = I
  405 CONTINUE
      do 44444 i=1,kdim4
      icon(i)=0
44444 continue
      write(93,44445)
44445 format(' EVENTS IN MORE THAN 5 CYCLES IF ANY'/) 
      IF (ISKIP.EQ.1) GO TO 480
      COUNT = 0
      ICORT = 0
      DO 470 I = 1,MMAX
  410     IF (IA.LE.MAXCYCc) GO TO 415
              WRITE (out(4),4000) MAXCYC
 4000         FORMAT (/1X, '*** NUMBER OF CYCLES HAS EXCEEDED THE',
     +         ' ALLOWED MAXIMUM OF', I5, '.',
     +         //1X, '**** EXECUTION TERMINATED.'/)
              GO TO 9999
  415     ICYC = 0
          ISUP = 0
          DO 420 J = 1,KDIM5
              WVEC(J) = 0.0
  420     CONTINUE
  425     K = I + 1
          IF (K.GT.MMAX) GO TO 470
          DO 465 J = K,MMAX
              IF (C(I,J).GE.(C(J,I)-TOL)) GO TO 465
              DO 430 T = 1,MMAX
                  TEMP = C(I,T)
                  C(I,T) = C(J,T)
                  C(J,T) = TEMP
  430         CONTINUE
              DO 435 T = 1,MMAX
                  TEMP = C(T,I)
                  C(T,I) = C(T,J)
                  C(T,J) = TEMP
  435         CONTINUE
              ITEMP = IPOS(I)
              IPOS(I) = IPOS(J)
              IPOS(J) = ITEMP
              LIMT = KDIM5 - 1
              DO 440 T = 1,LIMT
                  WVEC(T) = WVEC(T+1)
  440         CONTINUE
              WVEC(KDIM5) = FLOAT (IPOS(I))
              COUNT = COUNT + 1
              IF (COUNT.LT.ITER) GO TO 445
                  WRITE (out(4),4020) ITER
 4020             FORMAT (/1X, '*** NUMBER OF ALLOWED MATRIX TRANS',
     +             'FORMATIONS (ITER =', I6, ') EXCEEDED.',
     +             //1X, '**** EXECUTION TERMINATED IN M/PROG.'/)
                  GO TO 9999
  445         IF (IALPHA.EQ.0 .OR. IALPHA.EQ.1)  GO TO 450
                  WRITE (out(4),4050)
 4050             FORMAT (//20H SEQUENCING PROGRESS)
                  WRITE (out(4),4060) (IPOS(L), L = 1,MMAX)
 4060             FORMAT (1X, 20I5)
  450         CONTINUE
C
C          TEST AND CORRECT FOR CYCLICITY
C

              ISUP = ISUP + 1
              IF (ISUP.LE.100 .OR. WVEC(1).LE.0.0) GO TO 425
              DO 455 T = 4,KDIM5
                  IF (WVEC(T).EQ.WVEC(1)) GO TO 460
  455         CONTINUE
              GO TO 425
  460         ICYC = T - 1
      CALL CYCLE(ICORT,ICYC,WVEC,iqcyc,IA,A,B,CC,IPOS,out(4),isrt)
      do 30004 k=1,icyc
      kkk=iqcyc(k)
      icon(kkk)=icon(kkk)+1
30004 continue
         if(isrt.ne.0.and.isrt.ne.1) goto 461
              GO TO 410
  465     CONTINUE
  470 CONTINUE
      do 20005 i=1,kdim4
      if(icon(i).gt.5) write(93,20006)i,icon(i)
20005 continue
20006 format(2i4)
C
C  REPLACE ELEMENTS ZEROED IN CORRECTION OF CYCLICITY
C
      if(isrt.eq.0.or.isrt.eq.1) then
      ICORT = 1
      CALL CYCLE(ICORT,ICYC,WVEC,iqcyc,IA,A,B,CC,IPOS,out(4),isrt)
      endif
C
C
C  OUTPUT FINAL ORDER RELATION MATRIX, OPTIMUM SEQUENCE
C  AND RUN CONDITIONS
C
C
      do 10001 i=1,mmax
      do 10002 j=1,mmax
      if(j.lt.i.and.c(i,j).lt.0.5) goto 10002
      if(j.le.i) jvan(i,1)=j
      goto 10001
10002 continue
10001 continue
      do 10003 j=1,mmax
      do 10004 i=1,mmax
      if(j.lt.i.and.c(i,j).lt.0.5) goto 10004
      if(j.le.i) jvan(j,2)=i
10004 continue
10003 continue
      IF (IOMAT.NE.1) GO TO 480
          WRITE (out(4),4070)
 4070     FORMAT (///1X, 'FINAL ORDER RELATION MATRIX')
          DO 475 I = 1,MMAX
              WRITE (out(4),4080) I
 4080         FORMAT (//2X,I3)
              WRITE (out(4),3080) (C(I,J), J = 1,MMAX)
  475     CONTINUE
C
  480 WRITE (out(4),4090)
 4090 FORMAT (///// ' OPTIMUM SEQUENCE OBTAINED VIA RANKING'/)
      WRITE (out(4),4060) (IPOS(L), L = 1,MMAX)
      DO 485 I = 1,MMAX
          IRCODE(I) = ICODE(IPOS(I))
  485 CONTINUE
C
C  IPOS   - HOLDS THE OPT. SEQ.  (ENCODED RASC INDEX-NUMBERS)
C  IRCODE - HOLDS THE OPT. SEQ.  (ORIGINAL CODE NUMBERS)
C
      if(isrt.ne.0.and.isrt.ne.1) then
      do 491 i=1,mmax
      ircode2(i)=ircode(i)
  491 continue
      do 824 i=1,mmax-1
      j=i+1
      if(c(i,j).eq.c(j,i).and.c(i,j).gt.0.0) then
      ircode2(j)=-iabs(ircode2(j))
      endif
c      write(*,*) i,j,c(i,j),c(j,i),ircode2(i),ircode2(j)
  824 continue
      do 801 j=1,ib
      do 802 kk=1,kdim8
      iqcyc3(j,kk)=0
  802 continue
  801 continue
      do 803 j=1,ib
      do 804 kk=1,kdim5
      iqcyc3(j,kk)=iqcyc2(j,kk)
  804 continue
  803 continue
  811 do 805 i=1,ib-1
      do 806 j=i+1,ib
      do 821 ki=1,kdim8
      if(iqcyc3(j,ki).eq.iqcyc3(i,1)) then
      imemo=i
      do 807 k=ki+1,kdim8-1
      if(iqcyc3(i,k-ki+1).gt.0) then
      do 808 kk=1,kdim8-k
      kkk=kdim8+1-kk
      iqcyc3(j,kkk)=iqcyc3(j,kkk-1)
  808 continue
      iqcyc3(j,kkk-1)=iqcyc3(i,k-ki+1)
      endif
  807 continue
      goto 814
      endif
  821 continue
  806 continue
  805 continue
      goto 815
  814 ib=ib-1
      do 809 i=imemo,ib
      do 810 k=1,kdim8
      iqcyc3(i,k)=iqcyc3(i+1,k)
  810 continue
  809 continue
      goto 811
  815 do 812 i=1,ib
c      write(*,*) (iqcyc3(i,j),j=1,10)
  812 continue
      j=1
      do 816 i=1,mmax
      do 817 k=1,ib
      if(iabs(ircode2(i)).eq.iqcyc3(k,1)) then
      do 818 kk=1,kdim8
      iqcyc4(j,kk)=iqcyc3(k,kk)
  818 continue
      j=j+1
      endif
  817 continue
  816 continue
      do 819 i=1,ib
c      write(*,*) (iqcyc4(i,j),j=1,10)
  819 continue
      icount=0
      do 822 i=1,ib
      do 823 j=1,kdim8
      iqcyc3(i,j)=iqcyc4(i,j)
      if(j.gt.1.and.iqcyc3(i,j).gt.0) icount=icount+1
  823 continue
  822 continue
      ii=0
      j=1
      ib1=ib+1
      do 493 i=1,mmax
      if(j.gt.ib1) goto 492
      if(iabs(ircode2(i)).eq.iqcyc3(j,1)) then
      do 496 kk=1,kdim8
      if(iqcyc3(j,kk).eq.0) then
      j=j+1
      goto 493
      endif
      ii=ii+1
      ircode(ii)=iqcyc3(j,kk)
      if(kk.gt.1) ircode(ii)=-ircode(ii)
c      write(*,*) i,ii,ircode(ii)
  496 continue
      endif
      ii=ii+1
      ircode(ii)=ircode2(i)
c      write(*,*) i,ii,ircode(ii)
  493 continue
  492 continue
c      mmax=mfmax
c Note by FPA, 9 August, 1995. The preceding statement corresponds
c to mfmax=mmax (see before). Its purpose was to restore mmax to its
c original size as is now done in the next statement. However, although
c the preceding statement worked for one problem, it did not give good
c results for another problem because (arbitrarily??) mfmax (and
c therefore mmax) was set equal to 0. How could this happen?
      mmax=mmax+icount
      qdar(mmax+1)=qdar(mmax)+1.0
      endif
      WRITE (out(4),4100)
 4100 FORMAT (///// ' OPTIMUM SEQUENCE USING ORIGINAL CODE NUMBERS')
      WRITE (out(4),4110) (IRCODE(I), I = 1,MMAX)
 4110 FORMAT (1X, 20I5)
      if(isrt.ne.0.and.isrt.ne.1) goto 501
      WRITE (out(4),4130) COUNT, ITER
 4130 FORMAT (///// ' RANKING SOLUTION OBTAINED WITH:' //10X, I5,
     +  ' ITERATIONS OUT OF MAXIMUM',I7)
      WRITE (out(4),4150) CRIT1, TOL
 4150 FORMAT (//10X, 'CRITICAL TRANSPOSE ELEMENT SUM OF', F6.1,
     +  //10X, 'TOLERANCE OF', F6.1)
C
C
C     EVENT RANGES
C     ------------
C
C  PRINT SEQUENCE WITH NAMES (EVENT LABELS) AND RANGES
C  (WE MAKE USE OF THE CUM.ORDER MATRIX AND THE CORRESPONDING
C    POSITION MATRIX)
C
C
      DO 520 I = 1,MMAX
          K = I
  500     K = K - 1
              IF (K.EQ.0) GO TO 505
              ARG = C(K,I) - C(I,K)
              IF (ARG.LE.0.0) GO TO 500
  505     INK = K
          K = I
  510     K = K + 1
              IF (K.EQ.(MMAX+1)) GO TO 515
              ARG = C(I,K) - C(K,I)
              IF (ARG.LE.0.0) GO TO 510
  515     JNK = K
          IRANGE(I,1) = INK
          IRANGE(I,2) = JNK
  520 CONTINUE
C
      write (out(5),5008)
      write (out(5),5009)
      WRITE (OUT(5),5010)
      WRITE (OUT(5),5020)
      WRITE (OUT(5),5030)
      WRITE (OUT(5),5040)
 5008 format(//////7x,'RANKING ANALYSIS - PRINCIPAL RESULTS')
 5009 format(7x,'____________________________________')
 5010 FORMAT (////7x,'OPTIMUM SEQUENCE TABULATED WITH EVENT RANGES',
     +  ' AND LABELS:'/)
 5020 FORMAT (7X, 'SEQUENCE  EVENT   RANGE   EVENT')
 5030 FORMAT (7X, 'POSITION  NUMBER', 10X, 'NAME')
 5040 FORMAT (/)
      do 10530 i=1,1000
      iunc(i,1)=99
      iunc(i,2)=99
      nopt(i)=99
      sdopt(i)=9.999
10530 continue
      DO 530 I = 1,MMAX
          ID = IRCODE(I)
          iunc(id,1)=irange(i,1)-i+1
          iunc(id,2)=irange(i,2)-i-1
          JIRCOD(I) = ID
          WRITE (OUT(5),5050) I, ID, IRANGE(I,1), IRANGE(I,2),
     +      (ITITLE(ID,J), J = 1,10)
          WRITE (7,5047) I, ID, IRANGE(I,1), IRANGE(I,2),
     +      (ITITLE(ID,J), J = 1,10)
 5047     FORMAT (1X, I8, I9, I8, I4, 3X, 10A4)
 5050     FORMAT (1X, I8, I9, I8, '-', I3, 3X, 10A4)
  530 CONTINUE
      WRITE (OUT(5),5048)
 5048 FORMAT(//7X, 'NOTE: RANGES DEFINE OUTER LIMITS IN THE POSITION',
     +' SEQUENCE.',
     +/13X,'EVENTS CAN OCCUR ANYWHERE WITHIN THESE LIMITS.',
     +/13X,'THIS RANGE IS NOT STRATIGRAPHIC.')
C
C  JIRCOD  - A COPY OF THE OPTIMUM SEQUENCE OBTAINED BY RANKING.
C            (ORIGINAL FOSSIL NUMBERS).  TO BE USED AS INPUT TO WDIST().
C
  501 if(isrt.ne.0.and.isrt.ne.1.) then
      write(out(5),5011)
      write(out(5),5021)
      write(out(5),5031)
 5011 format(/////7x,'CONDENSED OPTIMUM SEQUENCE'/)
 5021 format(7x, 'SEQUENCE  FOSSIL    RANK     FOSSIL')
 5031 format(7x, 'POSITION  NUMBER  DISTANCE    NAME'/)
      ii=0
      do 502 i=1,mmax
      id=ircode(i)
      if(id.lt.0) then
      ii=ii-1
      qdar(i)=ii-1.0
      iqdar(i)=qdar(i)
      endif
      ii=ii+1
      qdar(i)=ii-1.0
      iqdar(i)=qdar(i)
      idid=iabs(id)
      write(out(5),5051) i,id,iqdar(i),(ititle(idid,j),j=1,10)
 5051 format(5x,i8,i9,i8,6x,10a4)
  502 continue
      write(out(5),5063)
      write(out(5),5061)
      write(out(5),5062)
 5063 format(//' EXPLANATORY NOTE:'/)
 5061 format(' A NEGATIVE FOSSIL NUMBER INDICATES THAT AN EVENT',
     +' IS (ON AVERAGE)'/,' COEVAL WITH THE EVENT PRECEDING IT IN THE',
     +' CONDENSED OPTIMUM SEQUENCE.'/)
 5062 format(' THE CONDENSED OPTIMUM SEQUENCE CONTAINS',
     +' CLUSTERS OF EVENTS WHICH ARE'/,' COEVAL ON THE AVERAGE.'
     +'  THESE CLUSTERS ARE COMPARABLE TO CLUSTERS'/,' OF EVENTS IN A',
     +' SCALED OPTIMUM SEQUENCE FOR VERY SMALL INTEREVENT DISTANCES.'/,
     +' CLUSTERS CONSIST OF CYCLES AND PAIRS OF EVENTS WITH',
     +' EQUAL (NON-ZERO)'/,' TRANSPOSE ELEMENTS.  CYCLES ARE OBTAINED',
     +' IN THE USUAL WAY BUT INSTEAD OF'/,' CONTINUING THE ALGORITHM',
     +' AFTER BREAKING EACH CYCLE AS BEFORE, A NEW'/,' COMPOSITE EVENT',
     +' IS DEFINED WHICH CONSISTS OF THE CYCLING EVENTS.'/,
     +' THIS IMPLIES THAT AN EVENT BELONGS TO AT MOST ONE CLUSTER',
     +' (IT CANNOT'/,' PARTICIPATE IN MORE THAN ONE CYCLE AS IN',
     +' THE MODIFIED HAY METHOD).'/)
     +
c      goto 9998
      endif
      if (itab1.eq.1.or.iscat.eq.1) then
      if(isrt.eq.0.or.isrt.eq.1) then
      if(itab1.eq.1.and.ns.lt.46) then
      write(out(6),9388)
 9388 format(///'THE FOLLOWING RESULTS CONTAIN OPTIMUM SEQUENCE AFTER'
     +' RANKING')
 9389 format(///'THE FOLLOWING RESULTS CONTAIN OPTIMUM SEQUENCE AFTER'
     +' SCALING')
      iprint=0
      call tab1 (out(6),iprint)
      iprint=1
      endif
      endif
C  Section from readin.for included here on 13 June, 1995
      write (out(6),9390)
 9390 format(///1x,'SEQUENCE OF WELLS:'//)
      do 9395 i=1,ns
      write (out(6), 9398) i, (name(i,j),j=1,10)
 9395 continue
 9398 format(1x,i4,3x,10a4)
      endif
      if(isrt.eq.0.or.isrt.eq.1) then
      if (iscore.eq.1) then
      rewind 23
      do 836 i=1,mmax
      write(23) (c(i,j),j=1,mmax)
  836 continue
      rewind 23
      nrpt=mod(ns,17)
      iprin=0
      call score (out(6),nrpt,itpt,iprin)
      iprin=1
      rewind 23
      do 837 i=1,mmax
      read(23) (c(i,j),j=1,mmax)
  837 continue
      rewind 23
      endif
      if(isrt.eq.0.or.isrt.eq.1) then
      do 279 i=1,mmax+1
        qdar(i)=i-1.0
  279 continue
      endif
c      IF (ISCAT.NE.1)  GO TO 534
      if(iscat.eq.1) then
      WRITE (64,5055)
      write(64,5045)
      endif
      endif
 5055     FORMAT (/////'CORRELATION OF WELL SEQUENCE DATA TO RANKED',
     +' OPTIMUM SEQUENCE')
 5045     format('_________________________________________________',
     +'___________'/)
 5056     FORMAT (/////'CORRELATION OF WELL SEQUENCE DATA TO SCALED',
     +' OPTIMUM SEQUENCE')
 5046     format('_________________________________________________',
     +'___________'/)
      if(isrt.ne.0.and.isrt.ne.1) then
      if(iscat.eq.1) write(64,5054)
 5054 format(/////'CORRELATION OF WELL SEQUENCE DATA TO CONDENSED',
     +' OPTIMUM SEQUENCE'/)
      do 825 i=1,mmax
      ircode(i)=iabs(ircode(i))
  825 continue
      endif
      ifit=0
  835 rewind 23
      do 536 i=1,mmax
      write(23) (c(i,j),j=1,mmax)
  536 continue
      rewind 23
      itsa=0
      if(ifit.eq.0) then
      write(63,5560)
      write(63,5561)
      write(72,10071) ialpha,ns
      endif
      if(ifit.eq.1) then
      write(66,5562)
      write(66,5561)
      endif
 5560 format('EVENT DEVIATIONS PER WELL FOR RANKING SOLUTION')
 5561 format('______________________________________________')
      do 53301 is=1,ns
      do 53302 i=1,kdim3
      iqd(i)=tmat(is,i)
      if(iqd(i).eq.0) goto 53302
      ned=i
53302 continue
      if(ifit.eq.0) write(72,50601) is,ned
      if(ifit.eq.1) write(75,50601) is,ned
50601 format(' WELL # ',i2,' WITH ',i3,' EVENTS')
c53303 continue
53301 continue
      DO 533 IS = 1,NS
              DO 531 I = 1,KDIM3
                  IQD(I) = TMAT(IS,I)
                  IF (IQD(I).EQ.0) GO TO 532
                  NED = I
  531         CONTINUE
  532 continue
      if(iscat.ne.1) goto 5341
      if(ifit.eq.0) WRITE (64,5060) (NAME(IS,J), J = 1,10)
      if(ifit.eq.1) write (65,5060) (name(is,j),j=1,10)
 5341 continue
      if(ifit.eq.0) then
 5562 format('EVENT DEVIATIONS PER WELL FOR SCALING SOLUTION')
      write(63,5060) (name(is,j),j=1,10)
      write(72,5060) (name(is,j),j=1,10)
      write(63,5059) is,ned
 5059 format (' WELL # ',i2,' WITH ',i3,' EVENTS'/)
      write(63,5058)
      write(72,5058)
      endif
      if(ifit.eq.1) then
c      write(66,5562)
c      write(66,5561)
      write(66,5060) (name(is,j),j=1,10)
      write(75,5060) (name(is,j),j=1,10)
      write(66,5059) is,ned
      write(66,5058)
      write(75,5058)
      endif
      if(icasc.ne.0) then
      write(25,5060) (name(is,j),j=1,10)
      write(25,5059) is,ned
      endif
      if(icasc.eq.1) then
      write(25,5069)
      write(25,5071)
      endif
c      endif
 5058 format(2x,'i',4x,'X(i)',4x,'YMAX-Y(i)',1x,'EXPECTED',2x,
     +'DEVIATION  NO.  NAME'/)
 5069 format(2x,'i',4x,'X(i)',4x,'EXPECTED',2x,'ERROR BAR')
 5071 format(27x,'INPUT'/)
 5060 FORMAT (///2X,10A4/)
      call fit(iqd,af1,bf1,cf1,ned,qdar,dev,sdev,icasc,ifit)
      sd(is)=sdev
      af(is)=af1
      bf(is)=bf1
      cf(is)=cf1
c      if (is.eq.1) then
c      write(*,*) ned,sdev
c      do 830 i=1,ned
c      write(*,*) i,iqd(i),qdar(i),dev(i)
c  830 continue
c      endif
      do 829 i=1,mmax
      bindev(i)=0.0
  829 continue
      do 826 k=1,mmax
      do 827 i=1,ned
      if(ircode(k).eq.iabs(iqd(i))) bindev(k)=dev(i)
  827 continue
c      if(ns.le.kdim2) then
      if(bindev(k).eq.0.0) bindev(k)=999.0
c      if(is.le.kdim2/2) tmat(is+kdim2,k)=1000.0*bindev(k)
c      if(is.gt.kdim2/2) ix(is+kdim2/2,k)=1000.0*bindev(k)
      c(is,k)=bindev(k)
c      endif
  826 continue
c      do 828 i=1,mmax
c      if(is.eq.1) write(*,*) i,ircode(i),bindev(i)
c  828 continue
      do 831 i=1,mmax
      ircodeo(i)=0
      ircodea(i)=0
  831 continue
      do 832 i=1,ned
      if(abs(dev(i)).gt.sdev.and.abs(dev(i)).lt.2.0*sdev) then
      do 833 k=1,mmax
      if(ircode(k).eq.iabs(iqd(i))) ircodeo(k)=ircode(k)
  833 continue
      endif
      if(abs(dev(i)).gt.2.0*sdev) then
      do 834 k=1,mmax
      if(ircode(k).eq.iabs(iqd(i))) ircodea(k)=ircode(k)
  834 continue
      endif
  832 continue
      if(iscat.ne.1) goto 533
      if(ifit.eq.0) CALL SCATTR (IQD,NED,IS,ircodeo,ircodea,out(7),ifit,
     +itsa)
      if(ifit.eq.1) call scattr(iqd,ned,is,ircodeo,ircodea,out(10),ifit,
     +itsa)
C Addition made on June 13, 1995
      if(ifit.eq.0) write (64,5901)
      if(ifit.eq.1) write (65,5901)
              do 901 i=1,ned
              id=iabs(iqd(i))
      if(ifit.eq.0) write(64,5900) (ititle(id,j),j=1,10),iqd(i)
      if(ifit.eq.1) write(65,5900) (ititle(id,j),j=1,10),iqd(i)
  901         continue
 5900 format(1x,10a4,i6)
 5901 format(//10x,'FOSSIL NAME',20x,'NUMBER'/)
  533     CONTINUE
      if(ifit.eq.0) write(63,5332)
      if(ifit.eq.1) write(66,5332)
      do 5331 is=1,ns
      if(ifit.eq.0) then
      write(63,5333) (name(is,j),j=1,10),sd(is)
      endif
      if(ifit.eq.1) then
      write(66,5333) (name(is,j),j=1,10),sd(is)
      endif
 5331 continue
 5332 format(//' STANDARD DEVIATION OF EVENTS PER WELL'/)
 5333 format(2x,10a4,f10.5)
c      if (ns.gt.kdim2) goto 838
c      do 836 i=1,10
c      write(*,*) ircode(i)
c      do 837 is=1,ns
c      if(is.le.kdim2/2) write(*,*) tmat(is+kdim2,i)
c      if(is.gt.kdim2/2) write(*,*) ix(is+kdim2/2,i)
c  837 continue
c  836 continue
      if(ifit.eq.0.and.ivar.eq.0) then
      if(isrt.ne.0.and.isrt.ne.1) goto 9998
      endif
      if(ifit.eq.0.and.ivar.eq.1) then
      write(75,10071) ialpha,ns
10071 format(//i2,'  TOTAL NO. OF WELLS = ',i3/)
      call deviat(sd,af,bf,cf,out(7),icasc,qdar,ifit,jvan,ivent,
     +nopt,sdopt,avesd)
      endif
      if(ifit.eq.0.and.ivar.eq.1) write(25,5334) ialpha,ift
      if(ifit.eq.1.and.ivar.eq.1) then
      call deviat(sd,af,bf,cf,out(7),icasc,qdar,ifit,jvan,ivent,
     +nopt,sdopt,avesd)
      write(25,5334) ialpha,ift
      endif
 5334 format(4x,i5,i2)
  838 if(ifit.eq.1) goto 800
      if(isrt.ne.0.and.isrt.ne.1) goto 9998
c Addition made on 21 June, 1995
c  534 continue
      if(iniq.eq.1) then
      oldmax=mmax
      kuniq=1
      do 280 i=1,ns
        rstep=0.0
        do 960 j=1,kdim3
          id=ix(i,j)
          if (id.gt.0) rstep=rstep+1.0
          if(id.eq.0) go to 965
  960   continue
  965   rstep=rstep-1.0
        do 970 k=1,kdim3
          rmat(k,1)=0.0
          rmat(k,2)=0.0
          rmat(k,3)=0.0
          ivec(k)=0
  970   continue
        icnt=0
        do 980 j=1,kdim3
          id=ix(i,j)
          if(id.eq.0) go to 990
          iad=iabs(id)
          icd=icode(iad)
          do 973 k=1,mmax
            if(ircode(k).eq.icd) go to 975
  973     continue
  975     id2=k
          if(id.lt.0) icd=icd*(-1)
          ivec(j)=icd
          rmat(j,1)=qdar(id2)
          icnt=icnt+1
  980   continue
  990   imax=id2
c
c    "icnt" is the length of the sequence
c
        if(icnt.ge.3) goto 992
        goto 280
  992   continue
c
        if(iniq.ne.1) goto 280
c insert for test on August 2nd (5 lines only)
c        if(nuniq(i).eq.1.and.i.eq.1) then
c        write(*,*) n,i,icnt
c        write(*,*) (ivec(j),j=1,icnt)
c        write(*,*) (rmat(j,1),j=1,icnt)
c        endif
        if(nuniq(i).eq.1) call xuniq2(n,i,icnt,ivec,rmat,runiq,kuniq)
  280 continue
      if(iniq.eq.0) goto 190
      write (out(5),7003)
      do 180 i=1,n
        if(runiq(i,2).le.0.0) goto 180
        runiq(i,1)=runiq(i,1)/runiq(i,2)
        mmax=mmax+1
        if(mmax.gt.kdim6+maxuq) write (out(5),2060)
 2060   format(///1x, '*** ERROR:  SUBSCRIPT OUT OF RANGE',
     +    ' AT ARRAY "QDAR"'/ 16x, 'FURTHER RESULTS MAY BE ERRONEOUS')
        qdar(mmax)=runiq(i,1)
        ircode(mmax)=i
  180 continue
  190 continue
      kuniq=0
      lll=1
      lout=0
c      write(out(11),7001)
 7001 format(//1x,'THE FOLLOWING OPTIMUM SEQUENCE WAS OBTAINED AFTER ',
     + 'RE-INSERTING THE')
c      write(out(11),7002)
 7002 format(' ','UNIQUE EVENTS')
 7003 format(/////1x,'POSITIONING OF UNIQUE EVENTS IN OPTIMUM SEQUENCE')
      call order (qdar,ipair,xlev,lll,lout,out(11))
        write(out(11),7004)
        write(81,7009) mmax,avesd
 7009 format(' TOTAL # OF ENTRIES = ',i3,';  AVE SD = ',f7.4)
        write(81,7004)
 7004 format(////1x,'FINAL OPTIMUM SEQUENCE WITH UNIQUE EVENTS'/)
        write(out(11),7005)
 7005 format(/4x,' RANK  EVENT #    FOSSIL EVENT NAME'/)
      nodon = .false.
      node = ipair(1,1)
        do 195 i=1,mmax
          logo = '  '
          if(iuniq(node,1).eq.1) logo = istar2
          if(iuniq(node,1).eq.1) nodon = .true.
          node = ipair(2,i)
          id=ircode(i)
          write(out(11),7006) i,id,logo,(ititle(id,j),j=1,10)
      write(81,7008)i,iunc(id,1),iunc(id,2),id,logo,(ititle(id,j),j=1,10
     +),nopt(id),sdopt(id)
 7006 format(1x,i7,i8,3x,1a3,10a4)
 7008 format(1x,i4,1x,i3,i2,i5,1a3,10a4,i4,f7.3)
  195 continue
      if(nodon) write(out(11),7007)
 7007 format(/5x,'** INDICATES A UNIQUE (RARE) EVENT'/)
      do 281 i=1,mmax+1
        qdar(i)=0
  281 continue
      if(ialpha.eq.1.and.iutem(1,1).gt.0) then
c      write(71,10070) mopti,mopti,mmax
      endif
      mmax=oldmax
      endif
      kuniq=0
c end of addition made on 21 June, 1995
      IF (IALPHA.EQ.1) GO TO 535
          GO TO 9998
C       --------------------------------------------
C
C     SCALING ANALYSIS
C     ----------------
C
C  SECOND MODIFICATION OF CUMULATIVE ORDER MATRIX
C
  535 rewind 23
      do 537 i=1,mmax
      read(23) (c(i,j),j=1,mmax)
  537 continue
      rewind 23
      IF (CRIT2.LE.CRIT1) GO TO 590
          IKNT = 0
          DO 560 I = 1,IMAX
              L = I + 1
              DO 550 J = L,MMAX
                  IF (C(I,J)+C(J,I).GE.CRIT2) GO TO 550
                  C(I,J) = 0.0
                  C(J,I) = 0.0
                  IKNT = IKNT + 1
  550         CONTINUE
  560     CONTINUE

          IF (IOMAT.NE.1) GO TO 590
          WRITE (out(4),3070)
          DO 570 I = 1,MMAX
              WRITE (out(4),3030)
              WRITE (out(4),3080) (C(I,J), J = 1,MMAX)
  570     CONTINUE
C
C     (IF LLL = 1,  THEN RESULTS OF SCALING WILL NOT BE PRINTED)
  590 LLL = INOSC
      IF (INOSC.EQ.1) LLL = 0
      out(8)=67
      out(10)=67
      if(inosc.ne.0) then
      write(61,5998)
      write(61,5999)
      endif
 5998 format(////7x,'SCALING ANALYSIS - PRINCIPAL RESULTS')
 5999 format(7x,'____________________________________')
      WRITE (out(8),6000)
      WRITE (out(8),6010)
      write (out(8),6011)
 6011 format(//'NOTE: INITIAL RESULTS ARE FOR SCALING WITHOUT FINAL',
     +/'REORDERING. HOWEVER, OCCURRENCE TABLE AND SUBSEQUENT RESULTS',
     +/'APPLY TO FINAL SCALING SOLUTION'//)
c      if(ialpha.eq.1.and.ivar.eq.1) then
c         write(63,6000)
c         write(63,6010)
c      endif
 6000 FORMAT (////5X,'SUPPLEMENTARY SCALING ANALYSIS RESULTS')
 6010 FORMAT (5X,'______________________________________')
      WRITE (out(8),5070) CRIT2, IKNT, MMSQ
 5070     FORMAT (//1X, 'SECOND MODIFICATION OF ORDER MATRIX:',
     +      /1X, 'BASED ON CRIT2 = ', F5.1, ', A TOTAL OF', I5,
     +      ' PAIRS (OUT OF', I5, ') HAVE BEEN ZEROED.')
C
C  EVALUATION OF OPTIMUM SEQUENCE BASED ON UNWEIGHTED AND WEIGHTED
C      DISTANCE ANALYSIS
C
c      WRITE (out(8),6020)
 6020 FORMAT (//' EVALUATION BASED ON UNWEIGHTED AND WEIGHTED',
     + ' DISTANCE ANALYSIS')
C
C  COMPUTE NORMAL Z VALUES OF FREQUENCIES CALCULATED FROM ORDER MATRIX
C
      CALL NORMZ (AAA,LLL,IOMAT,out(8))
C
C  COMPUTE 'DISTANCES' BETWEEN EVENTS AND CONSTRUCT DENDROGRAM
C
C
      do 7777 i=1,mmax
      ircode(i)=jircod(i)
 7777 continue
      CALL DIST (QDAR,MPAIR,AAA,LLL,out(8))
      LOUT = 0
      CALL ORDER (QDAR,IPAIR,XLEV,LLL,LOUT,out(8))
      IF (LLL.NE.0) WRITE (out(8),7010)
 7010 FORMAT (/////5X,'DENDROGRAM OF UNWEIGHTED INTEREVENT DISTANCES')
      IF (LLL.NE.0) CALL DENDRO (itony,iutem,IPAIR,XLEV,out(8),st)
C
C  REPEAT DISTANCE ANALYSIS WITH WEIGHTED DIFFERENCES
C
      LLL = INOSC
      CALL WDIST (JIRCOD,QDAR,MPAIR,AAA,LLL,INEG,CRIT2,out(8),st,ik)
      LOUT = 1
      CALL ORDER (QDAR,IPAIR,XLEV,LLL,LOUT,out(8))
      IF (LLL.NE.0) WRITE (out(8),7020)
 7020 FORMAT (////5X,'DENDROGRAM OF WEIGHTED INTEREVENT DISTANCES')
      IF (LLL.NE.0) CALL DENDRO (itony,iutem,IPAIR,XLEV,out(8),st)
C
C  REORDER FINAL RELATION MATRIX AND REPEAT CLUSTER ANALYSIS
C      FOR UNWEIGHTED AND WEIGHTED DIFFERENCES
C
      IF (IFIN.NE.1) GO TO 700
      if(inosc.ne.0) write(out(9),6901)
 6901 format(//' SCALING RESULTS OBTAINED AFTER FINAL REORDERING'/)
      DO 625 KKK = 1,5
          IRET  = 1
          DO 600 I = 1,MMAX
              IF (JIRCOD(I).NE.IRCODE(I)) IRET = 0
  600     CONTINUE
          IF (KKK.EQ.5) IRET = 1
          LLL = IRET
          IF (IRET.NE.1) GO TO 610
          WRITE (out(8),6030)
 6030     FORMAT (//)
          IF (IOMAT.NE.1) GO TO 620
              WRITE (out(8),6060)
 6060         FORMAT (////'  UPPER TRIANGLE OF NORMAL Z VALUES')
              DO 605 I = 1,MMAX
                  WRITE (out(8),6070)
 6070             FORMAT (///)
                  WRITE (out(8),6080) (C(I,J), J = 1,MMAX)
 6080             FORMAT (1X,15F8.3)
  605         CONTINUE
              GO TO 620
  610     CALL REORD (AAA,IPOS)
      do 20001 i=1,mmax
      do 20002 j=1,mmax
      if(j.lt.i.and.c(i,j).lt.0.5) goto 20002
      if(j.le.i) jvan(i,3)=j
      goto 20001
20002 continue
20001 continue
      do 20003 j=1,mmax
      do 20004 i=1,mmax
      if(j.lt.i.and.c(i,j).lt.0.5) goto 20004
      if(j.le.i) jvan(j,4)=i
20004 continue
20003 continue
              DO 615 I = 1,MMAX
                  JIRCOD(I) = IRCODE(I)
  615         CONTINUE
              LLL = 0
              IF (INOSC.GT.1 .AND. IRET.EQ.1) LLL = 1
C
C             COMPUTE NORMAL Z VALUES
C
              CALL NORMZ (AAA,LLL,IOMAT,out(8))
C
C             COMPUTE DISTANCES BETWEEN FOSSIL EVENTS
C
 620      LLL = 0
          IF (INOSC.GT.1 .AND. IRET.EQ.1) LLL = 1
          CALL DIST (QDAR,MPAIR,AAA,LLL,out(8))
          LOUT = 0
          CALL ORDER (QDAR,IPAIR,XLEV,LLL,LOUT,out(8))
          IF (LLL.NE.0) WRITE (out(8),7010)
          IF (LLL.NE.0) CALL DENDRO (itony,iutem,IPAIR,XLEV,out(8),st)
C
C      REPEAT DISTANCE CALCULATION WITH WEIGHTED DIFFERENCES
C
          LLL = 0
          IF (INOSC.GE.1 .AND. IRET.EQ.1) LLL = 1
          CALL WDIST (JIRCOD,QDAR,MPAIR,AAA,LLL,INEG,CRIT2,out(9),st,ik)
          LOUT = 1
          CALL ORDER (QDAR,IPAIR,XLEV,LLL,LOUT,out(9))
          IF (LLL.NE.0) WRITE (out(9),7020)
          IF (LLL.NE.0) CALL DENDRO (itony,iutem,IPAIR,XLEV,out(9),st)
          IF (IRET.EQ.1) GO TO 630
  625 CONTINUE
  630 if(inosc.ne.0) write(out(9),6900)
      if(inosc.ne.0) WRITE (out(9),6090) KKK
 6090 FORMAT (' REORDERING WITH SOLUTION AFTER ',I3,' ITERATIONS'///)
C
C
C
C     FINAL PROCESSING OPTIONS
C     ------------------------
C
C
C  CONSTRUCT OCCURRENCE TABLE FOR WELLS
C
  700 out(6)=67
      IF (ITAB1.EQ.1) then
      write(out(6),9389)
      CALL TAB1 (out(6),iprint)
      endif
C
C  STEP MODEL FOR INDIVIDUAL WELLS
C
      IF (ISCORE.EQ.1) THEN
      NRPT = MOD(NS,17)
      CALL SCORE (out(6),NRPT,itpt,iprin)
      ENDIF
C
c      IF (ISCAT.NE.1) GO TO 850
      if(iscat.eq.1) then
      WRITE (65,5056)
      write(65,5046)
      endif
      ifit=1
      goto 835
c      DO 800 I = 1,NS
c        DO 790 J = 1,KDIM3
c              IQD(J) = TMAT(I,J)
c              IF (IQD(J).EQ.0) GO TO 795
c              NED = J
c  790     CONTINUE
c  795     WRITE (OUT(10),5060) (NAME(I,J), J = 1,10)
c          CALL SCATTR (IQD,NED,I,ircodeo,ircodea,OUT(10))
  800 CONTINUE
C
C  PERFORM NORMALITY TEST ON INDIVIDUAL WELLS
C
c  850 continue
      IF (ICOMP.EQ.1) CALL COMP (QDAR,INIQ,OUT(10),kuniq,itnt)
C
C  ENTER UNIQUE EVENTS INTO SEQUENCE AND PRINT FINAL DENDROGRAM
C
      IF (INIQ.NE.1 .OR. ICOMP.NE.1) GO TO 9998
c          WRITE (OUT(11),6900)
          WRITE (OUT(11),7000)
 6900     FORMAT (//, ' PRECEDING RESULTS WERE OBTAINED AFTER FINAL')
 7000     FORMAT (//1X,'POSITIONING OF UNIQUE EVENTS IN FINAL SEQUENCE')
          CALL ORDER (QDAR,IPAIR,XLEV,LLL,0,OUT(11))
          write(out(11),7020)
          CALL DENDRO (itony,iutem,IPAIR,XLEV,OUT(11),st)

C THE LINES BETWEEM C ... AND C... WERE ADDED BY Z. HUANG IN JULY, 1992
C...
 9998     WRITE (61,7011)
          write(80,7011)
 7011 FORMAT(/////3X, 'SUMMARY OF DATA PROPERTIES AND RASC17 RESULTS:')
          WRITE (61,7021) N
          write(80,7021) n
 7021     FORMAT (/3X, 'NUMBER OF NAMES (TAXA) IN THE DICTIONARY ',
     +    '           ', I4)
          WRITE (61,7041) NS
          write(80,7041) ns
 7041     FORMAT (3X, 'NUMBER OF WELLS                          ',
     +    '           ', I4)
          WRITE (61,7060) N2
          write(80,7060) n2
 7060     FORMAT (3X, 'NUMBER OF DICTIONARY TAXA IN THE WELLS   ',
     +    '           ', I4)
          WRITE (61,7080) N1
          write(80,7080) n1
 7080     FORMAT (3X, 'NUMBER OF EVENT RECORDS IN THE WELLS     ',
     +    '           ', I4)
          if(isrt.eq.0.or.isrt.eq.1) WRITE (61,8000) (IA-1)/2
          if(isrt.eq.0.or.isrt.eq.1) write(80,8000) (ia-1)/2
          if(isrt.ne.0.and.isrt.ne.1) write(61,8001) ib
          if(isrt.ne.0.and.isrt.ne.1) write(80,8001) ib
 8000     FORMAT (3X, 'NUMBER OF CYCLES PRIOR TO RANKING        ',
     +    '           ', I4)
      xmopti=mopti
      crit4=.25*xmopti
      xia=(ia-1)/2
      if(xia.gt.crit4) then
      write(84,19998)
      write(87,19998)
      write(87,19999)
      endif
      write(84,19999)
19998 format('TYPE 1 WARNING: NUMBER OF CYCLES PRIOR TO RANKING'/'EXCEED
     +S 25% OF NUMBER OF EVENTS')
 8001     format (3x, 'NUMBER OF CLUSTERS WITH 3 OR MORE EVENTS ',
     +    '           ', i4)
          WRITE (61,8020) MOPTI
          write(80,8020) mopti
 8020     FORMAT (3X, 'NUMBER OF EVENTS IN THE OPTIMUM SEQUENCE ',
     +    '           ', I4)
          write(61,8030) ivent
          write(80,8030) ivent
 8030     format (3x, 'NUMBER OF EVENTS IN OPTIMUM SEQUENCE WITH',
     +    ' SD < ave SD', i3)
          IF (IALPHA.NE.0) then
          WRITE (61,8040)
          write(80,8040)
 8040     FORMAT (3X, 'NUMBER OF EVENTS IN THE FINAL SCALED OPTIMUM')
          WRITE (61,8060) MMAX
          write(80,8060) mmax
 8060     FORMAT (3X, '   SEQUENCE (INCLUDING UNIQUE EVENTS ',
     +    'SHOWN WITH **) ',I4)
          endif
          if(iscore.eq.1.and.ialpha.eq.0) write(61,8061)itpt
          if(iscore.eq.1.and.ialpha.eq.0) write(80,8061)itpt
          if(iscore.eq.1.and.ialpha.eq.1) write(61,8064)itpt
          if(iscore.eq.1.and.ialpha.eq.1) write(80,8064)itpt
          if(ialpha.eq.1.and.icomp.eq.1) write(61,8062)itnt
          if(ialpha.eq.1.and.icomp.eq.1) write(80,8062)itnt
          if(ialpha.eq.0.and.iscat.eq.1) write(61,8063)itsa
          if(ialpha.eq.0.and.iscat.eq.1) write(80,8063)itsa
          if(ialpha.eq.1.and.iscat.eq.1) write(61,8065)itsa
          if(ialpha.eq.1.and.iscat.eq.1) write(80,8065)itsa
          ns2=2.*ns
          if(itsa.gt.ns2) then
          write(85,20000)
          write(87,20000)
          write(87,19999)
          endif
20000 format('TYPE 2 WARNING: NUMBER OF AAAA EVENTS EXCEEDS'/'TWICE THE
     +NUMBER OF SECTIONS')
      if(ik.eq.1) write (87,29991) icrit
29991 format('TYPE 4 WARNING: SCALING OUTPUT CONTAINS 0.0000 INTEREVENT 
     +'/'DISTANCES BECAUSE THRESHOLD FOR MINUMUM NUMBER OF PAIRS OF EVEN
     +TS'/'CO-OCCURRING IN WELLS WAS SET AS LOW AS:',i3)
          write(85,19999)
          write(86,19999)
          write(87,19999)
19999 format(' ')
 8061 format(3x,'NUMBER OF STEPMODEL EVENTS WITH MORE THAN',
     +/'      SIX PENALTY POINTS AFTER RANKING   ',14x,I4)
 8064 format(3x,'NUMBER OF STEPMODEL EVENTS WITH MORE THAN',
     +/'      SIX PENALTY POINTS AFTER SCALING   ',14x,I4)
 8062 format(3x,'NUMBER OF NORMALITY TEST EVENTS SHOWN WITH * OR **',
     +'  ',I4)
 8063 format(3X,'NUMBER OF AAAA EVENTS IN RANKING SCATTERGRAMS',
     +'       ',I4)
 8065 format(3X,'NUMBER OF AAAA EVENTS IN SCALING SCATTERGRAMS',
     +'       ',I4)
          WRITE (61, 8080)
          write(80,8080)
          WRITE (61, 8080)
          WRITE (61, 8080)
 8080     FORMAT (3X, '    ')
C...
C
 9999 CLOSE (61)
c      rewind 75
c      do 101 i=1,mmax
c      read(75,1000) ircode(i), (ititle(ircode(i),j),j=1,10)
c      write(73,1000)ircode(i), (ititle(ircode(i),j),j=1,10)
c      read(75,2500)
c      write(73,2500)
c      do 102 k=1,ns
c      kk=k
c      read(75,2000) kk,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
c      write(73,2000)kk,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
c  102 continue
c      read(75,3000) nn,ave,sda,ska
c      write(73,3000)nn,ave,sda,ska
c      read(75,3500) ave
c      write(73,3500) ave
c      read(75,4000)
c      write(73,4000)
c      do 212 k=1,10
c      kk=k
c      read(75,4500) kk,xk1,xk2,ifreq(k),(ialpha(j),j=1,10)
c      write(73,4500)kk,xk1,xk2,ifreq(k),(ialpha(j),j=1,10)
c  212 continue
c  101 continue
      CLOSE (62)
      CLOSE (5)
c      CLOSE (10)
      CLOSE (99)
      CLOSE (7)
      close (22)
      close (23,status='delete')
      if(inew.eq.1) then
      close(10,status='delete')
      open(10,file=batfil,status='unknown')
      rewind 83
10112 read(83,10111) contents
      write(10,10111) contents
10111 format(a80)
      if(contents(1:4).eq.'LAST') goto 10103
      goto 10112
      endif
10103 close(83,status='delete')
      STOP
      END

      SUBROUTINE READIN (inew,iutem,OUT)
C
C ... SUBROUTINE TO READ IN ALL THE INPUT PARAMETERS, SEQUENCE                          
  
C  DATA, AND DICTIONARY.
C  EXECUTION IS TERMINATED IF ANY ERRORS ARE DISCOVERED.
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON  /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      COMMON  /DELTA/ IOCR, INIQ, CRIT1, TOL, AAA, CRIT2, MAX, ITER,
     +        IOMAT, ISRT, IALPHA, ITAB1, ISCORE, ICOMP, ISKIP, IFIN,
     +        INOSC, INEG, ISCAT, ivar, icasc, TMAT(KDIM1,KDIM3),ift
      COMMON  /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER item(maxuq),IuTEM(MAXUQ,2), TMAT, ITALLY(KDIM4), OUT(11)
      integer ievent(kdim3),iwell(kdim1),iwell2(kdim1)
      integer iutem2(kdim1,maxuq)
      real depth(kdim3),edepth(kdim3)
      CHARACTER*4 NAME, ITITLE, ITEMP
      character*3 answer
      character*1 blank,bla,blabla
      data blank/' '/
C
C  RETURNS THE OBSERVED DATA IN TMAT(*,*) VIA COMMON /DELTA/,
C          AND IN IX(*,*) VIA BLANK COMMON.
C  IUNIQ(KDIM4, 2)  IS A MAP SHOWING THE UNIQUE EVENTS IN COLUMN 1
C          AND THE MARKER HORIZONS IN COLUMN 2
C  NUNIQ(KDIM1) IS A MAP SHOWING WHICH SEQUENCES HAVE UNIQUE EVENTS
C  

C
C  READ THE RUN-PARAMETERS AND THE PROCESSING OPTIONS
C
ccccc      READ (5,1042) NS, IOCR, INIQ, ialpha,icasc, CRIT2
c      READ (5,1031) ITAPE, IOMAT, ISRT, IALPHA, ITAB1, ISCORE, ICOMP,
c     +      ISKIP, IFIN, INOSC, INEG, ISCAT, ivar, icasc
c      READ (5,1032) (OUT(I),I=1,11)
      iniq=0
      ialpha=0
      icasc=0
      read(5,1030) ns,iocr,iniq,ialpha,icasc,icrit,ift
      iniqi=1
      if(iniq.eq.0) then
      iniq=1
      iniqi=0
      endif
      inew=1
      read(10,5501)
      read(10,55500) bla
55500 format(a1)
 5501 format(1x)
      inew=1
      if(bla.eq.blank) inew=0
      rewind 10
      if(inew.eq.0) goto 500
      write(82,20000)
20000 format('DECIMAL DEPTH FILE')
      do 501 i=1,ns
      read (10,5000) (name(i,j),j=1,10)
      write(82,5000) (name(i,j),j=1,10)
      write(83,5000) (name(i,j),j=1,10)
      read(10,5001) bla,rth,wd
      blabla = bla
      if(bla.eq.'f'.or.bla.eq.'F')bla='m'
 5001 format(a1,f5.1,f7.1)
      write(82,5001) bla,rth,wd
      write(82,5002)
 5002 format('AUTHOR none')
      kindex=1
      kin=1
      read(10,5003) depth(1),ievent(1)
      edepth(1)=depth(1)
      if(blabla.eq.'f'.or.blabla.eq.'F') edepth(1)=.30480*edepth(1)
      do 502 k=2,kdim3
      kin=kin+1
      read(10,5003) depth(k),ievent(k)
      if(blabla.eq.'f'.or.blabla.eq.'F') depth(k)=.30480*depth(k)
 5003 format(f8.2,i4)
      ievent(k)=-ievent(k)
      if(depth(k).gt.depth(k-1)) then
      ievent(k)=-ievent(k)
      kindex=kindex+1
      edepth(kindex)=depth(k)
      endif
      if(depth(k).eq.0.0) then
      kind=(kindex+8)/9
      ievent(kin)=-999
      ki=(kin+19)/20
c      write(*,*)kind,ki
      do 503 kk=1,kind
      kkk=9*(kk-1)
      if (kk.lt.kind) write(82,5004) (edepth(kkk+ii),ii=1,9)
      if (kk.eq.kind) write(82,5004) (edepth(kkk+ii),ii=1,kindex-kkk)
 5004 format(f7.2,8(',',f7.2))
  503 continue
      do 504 kk=1,ki
      kkk=20*(kk-1)
      if(kk.lt.ki) write(83,5005) (ievent(kkk+ii),ii=1,20)
      if(kk.eq.ki) write(83,5005) (ievent(kkk+ii),ii=1,kin-kkk)
 5005 format(20i4)
  504 continue
      goto 501
      endif
  502 continue
  501 continue
      write(82,5006)
      write(83,5006)
 5006 format('LAST')
      rewind 82
      rewind 83
  500 crit2=icrit
c      if(iocr.gt.0) then
c      write(*,*)'    Do you wish to enter run parameters? (Y/N) '
c      read(*,2012) answer
c      if(answer.eq.'n'.or.answer.eq.'N') goto 8888
c      endif
c      write(*,*)'    Do you wish to revise the *.inp file? (Y/N) '
c      read(*,*) answer
      irev=0
c      if(answer.eq.'y'.or.answer.eq.'Y') irev=1
c      read(4,11111) irev
c11111 format(i1)
c      if(irev.eq.0) goto 8888
      goto 8888
      write(*,*)'    Enter total number of wells for this run: '
      read(*,*) nsnew
      if(nsnew.ne.ns) then
c      write(*,*)' '
      write(*,*) '    New number overrides previous number'
      endif
      ns=nsnew
c      write(*,*) ' '
      write(*,*)'    Enter minimum number of wells per event: '
      read(*,*) iocr
c      write(*,*)' '
      iniq=0
      ialpha=0
      icasc=0
      write(*,*)'    Will you use unique events/marker horizons? (Y/N) '
      read(*,2012) answer
 2012 format(a3)
      if(answer.eq.'y'.or.answer.eq.'Y') iniq=1
c      write(*,*)' '
      write(*,*)'    Do you wish to perform scaling? (Y/N) '
      read(*,*) answer
      if(answer.eq.'y'.or.answer.eq.'Y') ialpha=1
      if (ialpha.eq.1) then
c      write(*,*)' '
      write(*,*)'    Minimum number of wells for pairs of events: '
      read(*,*) crit2
      icrit=crit2
      endif
c      write(*,*)' '
      write(*,*)'    Do you wish to use CASC? (Y/N) '
      read(*,*) answer
      if(answer.eq.'y'.or.answer.eq.'Y') icasc=1
c      icodu1=0
c      icodu2=0
 8888 itape=1
      iter=90000
      crit1=1.0
      tol=0.0
      aaa=1.645
      iomat=0
      isrt=1
      itab1=1
      iscore=1
      icomp=1
      iskip=0
      ifin=1
      inosc=1
      ineg=1
      iscat=1
      ivar=1
      out(1)=1
      out(2)=0
      out(3)=0
      out(4)=0
      out(5)=1
      out(6)=0
      out(7)=0
      out(8)=0
      out(9)=1
      out(10)=0
      out(11)=1
 1030 format(7i2)
 1031 format(14i2)
 1032 format(11i2)
C
C ----- SET UP OUTPUT ---------------------
C
      DO 101 I=1,11
        IF (OUT(I).EQ.1) THEN
          OUT(I)=61
        ELSE
          OUT(I)=62
        ENDIF
 101  CONTINUE
C
C -----------------------------------------
C     
      
      WRITE (OUT(1),1040)
      WRITE (OUT(1),1041)
ccccc      rewind 5
      WRITE (OUT(1),1042) NS, IOCR, INIQi, Ialpha, icasc, icrit
ccccc      write (5,1042) ns,iocr,iniq,ialpha,icasc,crit2
ccccc      read (5,1042) ns,iocr,iniq,ialpha,icasc,crit2
c      WRITE (OUT(1),1045)
c      WRITE (OUT(1),1046)ITAPE,IOMAT,ISRT,IALPHA,ITAB1,ISCORE,ICOMP,
c     + ISKIP,IFIN
c      write (out(1), 2045)
c      write (out(1), 2046) INOSC, INEG, ISCAT,ivar,icasc
c 1040 FORMAT (///1X, 'VALUES OF INPUT PARAMETERS'//)
 1040 format(///1x)
 1041 FORMAT (1X, 'RUN PARAMETERS:    NS IOCR INIQ IALPHA ICASC ICRIT')
c     +  6X, 'AAA', 5X, 'CRIT2')
 1042 FORMAT (17X, I5, I5, I5, I7, i6, i6)
c 1045 FORMAT (/1X, 'PROCESSING OP:  ITAPE IOMAT ISRT IALPHA ITAB1',
c     +  ' ISCORE ICOMP ISKIP IFIN')
c 2045 format (17x, 'INOSC INEG ISCAT IVAR ICASC')
c 1046 FORMAT (18X, I4, I6, I5, I7, I6, I7, I6, I6, I5)
c 2046 format (16x, i6, i5, i6, i5, i6)
      DO 100 I = 1,KDIM4
          IRCODE(I) = 0
          IUNIQ(I,1) = 0
          IUNIQ(I,2) = 0
  100 CONTINUE
C
C  READ UNIQUE EVENTS AND MARKER HORIZONS
C
      if(iniqi.eq.1) READ (5,1050) (IuTEM(J,1), J = 1,20)
      if(iutem(1,1).eq.0) goto 110
 1050 FORMAT (20I4)
c      IF (IuTEM(1,1).EQ.0 .OR. INIQ.EQ.1)  GO TO 1055
c          WRITE (OUT(1),1051)
c 1051     FORMAT (///1X, '** WARNING:  UNIQUE EVENTS FOUND ON RECORD 2'
c     +      /15X, '"INIQ" SET EQUAL TO 1')
c          INIQ = 1
c 1055 continue
c      IF (INIQ.EQ.1) WRITE (OUT(1),1060)
 1060 FORMAT (///' RECORDS FOR THE FOLLOWING UNIQUE (RARE) EVENTS ',
     +'HAVE BEEN SELECTED:'/)
      if(irev.eq.1.and.iniq.eq.1) then
      write(*,*)'    Do you want new unique event numbers? (Y/N) '
      read(*,*) answer
      if(answer.eq.'y'.or.answer.eq.'Y') then
c      icodu1=1
      write(*,*)'    Enter up to 20 numbers in i4 format:'
      read(*,1050) (iutem(j,1),j=1,20)
      endif
      endif
      DO 105 I = 1,20
          ID = IuTEM(I,1)
          IF (ID.EQ.0) GO TO 110
          IUNIQ(ID,1) = 1
c          WRITE (OUT(1),1070) ID
 1070     FORMAT (5X,I4)
  105 CONTINUE
  110 if(iniqi.eq.1) READ (5,1050) (IuTEM(J,2), J = 1,20)
c      IF (IuTEM(1,2).EQ.0 .OR. INIQ.EQ.1)  GO TO 112
c          WRITE (OUT(1),1075)
c 1075     FORMAT (///1X, '** WARNING:  MARKER HORIZONS FOUND ON RECORD',
c     +      ' 3'/ 15X, '"INIQ" SET EQUAL TO 1')
c          INIQ = 1
c  112 continue
      if(irev.eq.1.and.iniq.eq.1) then
      write(*,*)'    Do you want new marker horizons? (Y/N) '
      read(*,*) answer
      if(answer.eq.'y'.or.answer.eq.'Y') then
c      icodu2=1
      write(*,*)'    Enter up to 20 numbers in i4 format:'
      read(*,1050) (iutem(j,2),j=1,20)
      endif
      endif
c      IF (INIQ.EQ.1) WRITE (OUT(1),1080)
 1080 FORMAT (///' THE FOLLOWING MARKER HORIZONS HAVE BEEN SELECTED:'/)
      DO 115 I = 1,20
          ID = IuTEM(I,2)
          IF (ID.EQ.0) GO TO 120
          IUNIQ(ID,2) = 1
c          WRITE (OUT(1),1070) ID
  115 CONTINUE
C
C  PERFORM DATA VALIDATION ON INPUT PARAMETERS
C
  120 CALL CHECK (NS,CRIT1,CRIT2,IOCR,IER)
      IF (IER.NE.0)  STOP
      if(irev.eq.1) then
      rewind 5
      write (5,1030) ns,iocr,iniq,ialpha,icasc,icrit
      write(5,1050) (iutem(j,1),j=1,20)
      write(5,1050) (iutem(j,2),j=1,20)
      endif
C
C  READ IN THE WELL DATA
C
      DO 130 I = 1,NS
          NUNIQ(I) = 0
          DO 125 J = 1,KDIM3
              TMAT(I,J) = 0
              IX(I,J) = 0
  125     CONTINUE
  130 CONTINUE
      IODATA = 5
      IF (ITAPE.EQ.1) IODATA = 10
      if(inew.eq.1) iodata=83
      DO 135 I = 1,NS
          READ (IODATA,1090) (NAME(I,J), J = 1,10)
 1090     FORMAT (10A4)
          K = 0
          kk=0
  131     READ (IODATA,1100) (ITEM(J), J = 1,20)
 1100         FORMAT (20I4)
              DO 133 J = 1,20
                  ITEMJ = ITEM(J)
                  IF (ITEMJ.EQ.-999)  GO TO 134
                  IF (ITEMJ.EQ.0)  GO TO 133
                  do 1101 ii=1,maxuq
                  if(iabs(itemj).eq.iutem(ii,1)) then
                  kk=kk+1
                  iutem2(i,kk)=itemj
c                  write(*,*)iutem2(i,ii)
                  endif
 1101             continue
                      K = K + 1
                      iwell(i)=k
                      iwell2(i)=kk
                      IF (K.LE.KDIM3)  GO TO 132
                          WRITE (OUT(1),1105)
 1105                     FORMAT (//1X, '*** ERROR:  SEQUENCE TOO LONG',
     +                        /12X, 'POSSIBLE CAUSE:  MISSING -999',
     +                        /12X, 'OTHERWISE:  USE A LARGER VERSION',
     +                        ' OF RASC'//' **** EXECUTION TERMINATED'/)
                          STOP
  132                 TMAT(I,K) = ITEMJ
  133         CONTINUE
              GO TO 131
  134     IF (K.LT.KDIM3)  TMAT(I,K+1) = 0
  135 CONTINUE
C
C  READ IN THE DICTIONARY
C
      I = 1
      READ (99,5000) (ITITLE(I,J), J = 1,10)
 5000 FORMAT (10A4)
 5500 IF (ITITLE(I,1).EQ.'LAST'.OR.ITITLE(I,1).EQ.'last'
     + .OR. I.EQ.KDIM4)  GO TO 6000
          I = I + 1
          READ (99,5000) (ITITLE(I,J), J = 1,10)
          GO TO 5500
 6000 IF (ITITLE(I,1).EQ.'LAST'.OR.ITITLE(I,1).EQ.'last')  GO TO 7000
          I = I + 1
          READ (99,5000) ITEMP
          IF (ITEMP.EQ.'LAST')  GO TO 7000
              WRITE (OUT(1),6200) KDIM4, KDIM4
 6200         FORMAT (///1X, '*** ERROR:  MORE THAN',I5,' EVENT',
     +         ' LABELS FOUND IN DICTIONARY'/16X, 'FOR DICTIONARIES',
     +         ' WITH MORE THAN',I5, ' NAMES,',/16X, 'PLEASE USE',
     +         ' A LARGER VERSION OF RASC')
              WRITE (OUT(1),6250)
 6250         FORMAT (///' **** EXECUTION TERMINATED IN',
     +          '   SUBROUTINE READIN')
              STOP
 7000 N = I - 1
C      WRITE (OUT(1),7500) N
C 7500 FORMAT (/1X, 'DICTIONARY:', I5, ' NAMES WERE READ IN.')
C
C  CHECK THE WELL DATA
C
      IERR = 0
      IERR2 = 0
      DO 9000 I = 1,NS
          DO 8300 J = 1,N
              ITALLY(J) = 0
 8300     CONTINUE
          DO 8500 J = 1,KDIM3
              ID = IABS (TMAT(I,J))
              IF (ID.EQ.0)  GO TO 9000
              IF (ID.GT.N)  GO TO 8370
              IF (ITALLY(ID).EQ.0)  GO TO 8400
                  WRITE (OUT(1),8340) (NAME(I,K), K = 1,10), ID
 8340             FORMAT (/' IN ',10A4,' FOSSIL',I5,' OCCURS MORE THAN',
     +              ' ONCE')
                  IERR = IERR + 1
                  GO TO 8500
 8370         WRITE (OUT(1),8380) (NAME(I,K), K = 1,5), ID
 8380             FORMAT (/' IN ',10A4,' FOSSIL',I5,
     +              '  -  EXCEEDS LIMIT OF YOUR DICTIONARY')
                  IERR2 = IERR2 + 1
                  GO TO 8500
 8400         ITALLY(ID) = 1
 8500     CONTINUE
 9000 CONTINUE
      IF (IERR.EQ.0)  GO TO 9200
          WRITE (OUT(1),9100) IERR
 9100     FORMAT (//' *** ERROR -  ON',I6,' OCCASION(S), AN EVENT',
     +      ' OCCURRED MORE THAN ONCE IN THE SAME WELL.'/15X, 'YOU',
     +      ' MAY NOT REPORT MORE THAN ONE OCCURRENCE OF ANY FOSSIL',
     +      ' IN ANY WELL.')
 9200 IF (IERR2 .EQ. 0) GO TO 9300
          WRITE (OUT(1),9220) IERR2, N
 9220     FORMAT (//' *** ERROR - ON',I6,' OCCASION(S), THERE',
     +      ' WAS A FOSSIL NUMBER GREATER THAN THE LARGEST',
     +      ' DICTIONARY NUMBER:',I5)
C
C  PRINT ORIGINAL SEQUENCE DATA

 9300 DO 9340 I = 1,NS
          DO 9320 J = 1,KDIM3
              IX(I,J) = TMAT(I,J)
 9320     CONTINUE
 9340 CONTINUE
      WRITE (OUT(1),9390)
 9390 FORMAT (///1x,'SEQUENCE OF WELLS:'//)     
      write(89,10001) ns
      if(ift.eq.0)write(89,10002)
      if(ift.eq.1)write(89,10003)
10001 format(' NUMBER OF WELLS =',i4)
10002 format(' DEPTHS IN METERS'/)
10003 format(' DEPTHS IN FEET'/)
      DO 9395 I = 1,NS
      WRITE (OUT(1), 10004) I, (NAME(I,J),J=1,10)
      last=iwell2(i)
      write(89,10005) i,(name(i,j),j=1,10),iwell(i),last
 9395 CONTINUE
      write (89,10006)
      do 9396 i=1,ns
      last=iwell2(i)
      write(89,10007) i,last,(iabs(iutem2(i,j)),j=1,last)
 9396 continue
10004 FORMAT(1X,I4,3X,10A4)
10005 FORMAT(1X,I4,3X,10A4,' ## of Events and UEs = ',2i3)
10006 format(' '/)
10007 format(1x,i4,3x,'The',i2,' UEs are ',20i5)
C     CALL ECHO (OUT(1))
C
      IF (IERR.EQ.0 .AND. IERR2.EQ.0)  GO TO 9400
          WRITE (OUT(1),6250)
          STOP
C
 9400 RETURN
      END
      SUBROUTINE SCORE (UNIT,NRPT,itpt,iprin)
C
C ... SUBROUTINE TO CALCULATE STEP MODEL SCORES FOR INDIVIDUAL WELLS
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      REAL    sum(kdim1),xk(kdim1),tau(kdim1)
      CHARACTER*4  NAME, ITITLE
      CHARACTER*30 FMT1
      character*5 lne(27),line(50)
      INTEGER UNIT,NRPT
      data lne/5h  0.0,5h  0.5,5h  1.0,5h  1.5,5h  2.0,5h  2.5,5h  3.0,
     +5h  3.5,5h  4.0,5h  4.5,5h  5.0,5h  5.5,5h  6.0,5h  6.5,5h  7.0,
     +5h  7.5,5h  8.0,5h  8.5,5h  9.0,5h  9.5,5h 10.0,5h 10.5,5h 11.0,
     +5h 11.5,5h 12.0,5h  AAA,5h     /
C
      isave=unit
      write(91,10001) mmax
10001 format(//' NUMBER OF EVENTS = 'i5)
10000 irepeat=0
      itpt=0
      mao=0
  200 continue
      irepeat=irepeat+1
      DO 20 I = 1,MMAX
          DO 10 J = 1,NS
              c(I,J) = -.000001
              tau(j)=0.0
              xk(j)=0.0
              sum(j)=0.0
   10     CONTINUE
   20 CONTINUE
cc Addition inserted on 5 May, 1997
      do 153 i=1,ns
      do 154 j=1,mmax
      ioc=ix(i,j)
      if(ioc.eq.0) goto 153
      iad=iabs(ioc)
      id=icode(iad)
      do 155 k=1,mmax
      if(id.ne.ircode(k)) goto 155
      c(k,i)=0.00
      goto 154
  155 continue
  154 continue
  153 continue
c      itpt=0
c      mao=0
C
      DO 150 I = 1,NS
          DO 140 J = 1,MMAX
              ID = IX(I,J)
              IF (ID.EQ.0) GO TO 150
              xk(i)=xk(i)+1.0
              IAD = IABS (ID)
              IELMT = ICODE(IAD)
C
              DO 30 K = 1,MMAX
                  IF (IRCODE(K).EQ.IELMT) GO TO 40
   30         CONTINUE
   40         IOP = K
C
              K = J
              IF (ID.LT.0) ITEST = 1
   50         CONTINUE
                  K = K - 1
                  IF (K.LT.1) GO TO 90
                  ID2 = IX(I,K)
                  IF (ITEST.EQ.1) GO TO 80
                  IAD2 = IABS (ID2)
                  IELMT2 = ICODE(IAD2)
                  DO 60 L = 1,MMAX
                      IF (IRCODE(L).EQ.IELMT2) GO TO 70
   60             CONTINUE
   70             IOP2 = L
                  IF (IOP2.LT.IOP) GO TO 50
                  if(irepeat.lt.3) c(IOP,I) = c(IOP,I) + 1.0
                  GO TO 50
   80             if(irepeat.eq.1) c(IOP,I) = c(IOP,I) + 0.5
                  IF (ID2.GT.0) ITEST = 0
                  GO TO 50
C
   90         K = J
              ITEST = 1
  100         K = K + 1
              IF (K.GT.KDIM3)  GO TO 140
              ID2 = IX(I,K)
              IF (ID2.EQ.0) GO TO 140
              IAD2 = IABS (ID2)
              IELMT2 = ICODE(IAD2)
              IF (ITEST.NE.1) GO TO 110
              if(irepeat.eq.1.and.ID2.LT.0) c(IOP,I) = c(IOP,I) + 0.5
              IF (ID2.GT.0) ITEST = 0
              IF (ID2.LT.0) GO TO 100
  110         CONTINUE
              DO 120 L = 1,MMAX
                  IF (IRCODE(L).EQ.IELMT2) GO TO 130
  120         CONTINUE
  130         IOP2 = L
              IF (IOP2.GT.IOP) GO TO 100
              if(irepeat.ne.2) c(IOP,I) = c(IOP,I) + 1.0
              GO TO 100
  140     CONTINUE
  150 CONTINUE
      if(irepeat.eq.1) then
      do 151 i=1,ns
      do 152 k=1,mmax
          if(c(k,i).gt.6.0) itpt=itpt+1
          if(c(k,i).gt.12.0) mao=1
  152 continue
  151 continue
      endif
C 
C     
      N1 = INT(NS/17)
 
      IF (N1.LT.1) THEN
      NSTOP=NS
      ELSE
      NSTOP=17
      ENDIF
      if(irepeat.eq.1) then
      WRITE (UNIT,1010)
      write(unit,1011)
      write(unit,1014)
      write(unit,1016)
      if (mao.eq.1) write(unit,1017)
      endif
      if(irepeat.eq.2) write (unit,1018)
      if(irepeat.eq.3) write (unit,1019)
      write(unit,1012)
      if(iprin.ne.1) write(unit,1013)
      if(iprin.eq.1) write(unit,10130)
 1010 FORMAT (//// ' STEP MODEL')
 1018 format (//// ' PENALTY POINTS FOR EVENTS LOWER THAN EXPECTED')
 1019 format (//// ' PENALTY POINTS FOR EVENTS HIGHER THAN EXPECTED')
 1011 format (' __________')
 1014 format(/'THE STEPMODEL GIVES A PENALTY POINT FOR EACH POSITION'
     c/'AN EVENT RECORD IN A WELL IS OUT OF PLACE, RELATIVE TO')
 1016 format('THE EVENT ORDER IN THE (SCALED) OPTIMUM SEQUENCE.')
 1017 format('AAA INDICATES MORE THAN 12 PENALTY POINTS.')
 1012 format(//11X, 'NAME', 23X,
     +  'NUMBER', 5X, 'WELL NUMBER')
 1013 format(3x,'(OPTIMUM SEQUENCE)')
10130 format(1x,'(SCALED OPTIMUM SEQUENCE)')
      WRITE (UNIT,1030) (I, I = 1,NSTOP)
 1030 FORMAT (1X,T48,17I5)
      WRITE (UNIT,1040)
 1040 FORMAT (/)
      DO 160 I = 1,MMAX
          ID = IRCODE(I)
          do 161 k=1,nstop
          if(c(i,k).lt.0.0) line(k)=lne(27)
          if(c(i,k).ge.0.0.and.c(i,k).le.12.0) then
          xl=2.0*c(i,k)+1.0
          idnew=int(xl)
          line(k)=lne(idnew)
          endif
          if(c(i,k).gt.12.0) line(k)=lne(26)
  161 continue
          WRITE (UNIT,1051) (ITITLE(ID,J), J = 1,10), ID,
     +                      (line(k), K = 1,NSTOP)
c 1050     FORMAT (1X, 10A4, I3, T48, 17F5.1)
 1051     format(1x,10a4,i3,t48,17a5)
  160 CONTINUE
c  Inserted on 14 June, 1995
C
C  Kendall's tau statistic
C
      if(irepeat.eq.1) then
      do 180 k=1,ns
          do 190 i=1,mmax
c              if(c(i,k).lt.0.0) c(i,k)=0.0
              sum(k)=sum(k)+c(i,k)
  190     continue
          tau(k)=(sum(k)*2.0)/(xk(k)*(xk(k)-1.0))
          tau(k)=1.0-tau(k)
  180 continue
      if(ns.lt.18) then
      write(unit,1061) (tau(k),k=1,ns)
      write(unit,1064)
      write(unit,1065)
      endif
      endif
 1061 format(//1x,'   KENDALL`S TAU =',29x,17f5.2//)
      IF (N1.GT.1)THEN
      WRITE (UNIT,1015)
 1015 FORMAT (//// ' STEP MODEL (CONTINUED)', ///11X, 
     +       'NAME', 23X, 'NUMBER', 5X, 'WELL NUMBER')
      WRITE (UNIT,1030) (I, I = 18,34)
      WRITE (UNIT,1040)
      DO 165 I = 1,MMAX
          ID = IRCODE(I)
          do 162 k=18,34
          if(c(i,k).lt.0.0) line(k)=lne(27)
          if(c(i,k).ge.0.0.and.c(i,k).le.12.0) then
          xl=2.0*c(i,k)+1.0
          idnew=int(xl)
          line(k)=lne(idnew)
          endif
          if(c(i,k).gt.12.0) line(k)=lne(26)
  162 continue
          WRITE (UNIT,1051) (ITITLE(ID,J), J = 1,10), ID,
     +                      (line(k), K = 18,34)
  165 CONTINUE
      ENDIF
C
C  The following statement was changed by FPA on 1/2/95
C
C     IF (NRPT.GT.0)THEN
      IF (NRPT.GT.0.AND.N1.GT.0)THEN
      WRITE (UNIT,1015)
      WRITE (FMT1,1035) NRPT
 1035 FORMAT('(1X, T48,',I2,'I5)')
      WRITE (UNIT,FMT1) (I, I = N1*17+1,NS) 
      WRITE (UNIT,1040)
c      WRITE (FMT2,1055) NRPT
c 1055 FORMAT('(1X, 10A4, I3, T48,',I2,'f5.1)')
      n17=n1*17+1
      DO 170 I = 1,MMAX
          ID = IRCODE(I)
          do 163 k=n17,ns
          if(c(i,k).lt.0.0) line(k)=lne(27)
          if(c(i,k).ge.0.0.and.c(i,k).le.12.0) then
          xl=2.0*c(i,k)+1.0
          idnew=int(xl)
          line(k)=lne(idnew)
          endif
          if(c(i,k).gt.12.0) line(k)=lne(26)
  163 continue
          WRITE (UNIT,1051) (ITITLE(ID,J), J = 1,10), ID,
     +                      (line(k), K = N17,NS)
  170 CONTINUE
      ENDIF
      if(ns.gt.17.and.irepeat.eq.1) then
      write (unit,1063)
      do 171 k=1,ns
        write(unit,1062) k,tau(k)
  171 continue
      write (unit,1064)
      write (unit,1065)
      endif
      if(irepeat.lt.3) goto 200
      if(unit.ne.91) then
      unit=91
      goto 10000
      endif
      unit=isave
 1063 format(//'  WELL NO.    KENDALL`S TAU'/)
 1062 format(6x,i4,12x,f5.3)
 1064 format(//'NOTE: KENDALL`S TAU IS A RANK CORRELATION COEFFICIENT.',
     +/'LIKE THE ORDINARY (PEARSON`S) PRODUCT-MOMENT CORRELATION',
     +/'COEFFICIENT, TAU VARIES BETWEEN 0 FOR COMPLETE LACK OF')
 1065 format('CORRELATION AND +1 OR -1 FOR MAXIMUM POSITIVE OR NEGATIVE'
     +,/'CORRELATION. TAU IS CALCULATED FROM THE PENALTY SCORES FOR',
     +/'EACH WELL AND EXPRESSES DEGREE OF CORRELATION WITH THE',
     +/'OPTIMUM SEQUENCE.'//)
      RETURN
      END
      SUBROUTINE TAB1 (UNIT,iprint)
C
C ... SUBROUTINE TAB1 CONSTRUCTS AN OCCURRENCE TABLE SHOWING THE
C  OCCURRENCE OF OPTIMUM SEQUENCE EVENTS IN EACH WELL
C
C  CALLED IN MAIN PROGRAMME.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      CHARACTER*2  IBLANK, IEX, ITAB(KDIM6,45), IWELL(KDIM1),
     +             IWELL1(KDIM1), IWELL2(KDIM1)  
      CHARACTER*4  NAME, ITITLE
      INTEGER UNIT
      DATA    IBLANK/'  '/, IEX/' X'/
C
C  MATRIX ITAB(45, KDIM1) IS THE OCCURRENCE TABLE
C
      isave=unit
      write(90,10001) mmax
10001 format(//' NUMBER OF EVENTS = ',i5)
10000 DO 20 I = 1,MMAX
          DO 10 J = 1,NS
              ITAB(I,J) = IBLANK
   10     CONTINUE
   20 CONTINUE
      DO 50 I = 1,NS
          DO 40 J = 1,KDIM3
              IOC = IX(I,J)
              IF (IOC.EQ.0) GO TO 50
              IAD = IABS (IOC)
              ID = ICODE(IAD)
              DO 30 K = 1,MMAX
                  IF (ID.NE.IRCODE(K)) GO TO 30
                  ITAB(K,I) = IEX
                  GO TO 40
   30         CONTINUE
   40     CONTINUE
   50 CONTINUE
      DO 55 I = 1, NS   
      WRITE(IWELL(I),'(I2)') I
      IWELL1(I)=IWELL(I)(1:1)
      IWELL2(I)=IWELL(I)(2:2)
   55 CONTINUE
      WRITE (UNIT,6000)
      write(unit,6001)
 6000 FORMAT (//// '   OCCURRENCE TABLE')
 6001 format('  __________________'//)
      WRITE (UNIT,2000)
      if(iprint.lt.1) write(unit,2001)
      if(iprint.eq.1) write(unit,20010)
 2000 FORMAT (11X, 'NAME', 26X, 'NUMBER', 3X, 'WELL NUMBER')
 2001 format(3x,'(OPTIMUM SEQUENCE)')
20010 format(1x,'(SCALED OPTIMUM SEQUENCE)')
      WRITE (UNIT,3000) (IWELL1(I), I = 1,NS)
      WRITE (UNIT,3000) (IWELL2(I), I = 1,NS)
 3000 FORMAT (1X,T51,45A2)
      WRITE (UNIT,4000)
 4000 FORMAT (' '/)
      DO 60 I = 1,MMAX
          ID = IRCODE(I)
          WRITE (UNIT,5000) (ITITLE(ID,J), J = 1,10), ID,
     +                      (ITAB(I,K),K = 1,NS)
 5000     FORMAT (1X,10A4,I6,T50, 45A2)
   60 CONTINUE
      if(unit.ne.90) then
      unit=90
      goto 10000
      endif
      unit=isave
      RETURN
      END
      SUBROUTINE WDIST (JIRCOD,QDAR,MPAIR,AAA,LLL,INEG,CRIT2,UNIT,st,ik)
C
C ... SUBROUTINE TO CALCULATE INTER-EVENT 'DISTANCES' FOR
C  WEIGHTED DIFFERENCES
C
C  ACCEPTS:   MPAIR  - AS GENERATED BY DIST()
C             JIRCOD - A COPY OF THE OPTIMUM SEQUENCE (ORIGINAL CODE
C                      NUMBERS) OBTAINED BY THE PREVIOUS SCALED SOLU-
C                      TION FROM THE MAIN PROGRAMME.
C                    THE FIRST TIME THROUGH, JIRCOD WILL CONTAIN A
C                      COPY OF THE OPTIMUM SEQUENCE FROM THE RANKING
C                      SOLUTION.
C             IUNIQ(*,2)  -  A MAP SHOWING THE UNIQUE EVENTS IN COLUMN
C                      1 OF IUNIQ AND THE MARKER HORIZONS IN COLUMN 2.
C
C  RETURNS:   QDAR   - FOR ORDER()
C             IRCODE - ANOTHER COPY OF THE OPTIMUM SEQUENCE OBTAINED
C                      IN THE PREVIOUS ITERATION.  (THIS COPY IS SENT
C                      TO ORDER().)
C
C  CALLED IN MAIN PROGRAMME.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +       C(KDIM6,KDIM6)
      COMMON /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      INTEGER UNIT
      DIMENSION JIRCOD(KDIM4), MPAIR(KDIM6+MAXUQ), QDAR(KDIM6+MAXUQ)
      dimension st(kdim6+maxuq)
      DATA  PI/3.141592654/, CONST3/-111.1111/, CONST4/999.9999/
C
      QDAR(1) = 0.0
      NMAX = MMAX - 1
      IF (LLL.EQ.0) GO TO 10
      WRITE (UNIT,1010)
 1010 FORMAT (//' DISTANCE ANALYSIS WITH',
     +' WEIGHTED DIFFERENCES'/)
      WRITE (UNIT,1020)
 1020 FORMAT (1X, 'RANK   EVENT    INTEREVENT  CUMULATIVE',
     + 3X, 'SUM DIFF  SAMPLE  WGT     S.D.' /8X, 'PAIRS     DISTANCE',
     + '    DISTANCE', 4X, 'Z VALUES   SIZE'/)
   10 DO 20 I = 1,MMAX
          IRCODE(I) = JIRCOD(I)
   20 CONTINUE
C
C  IT IS STILL ASSUMED THAT  C(I1,I2)  IS A "Z"-VALUE
C  WHENEVER  I1 .LT. I2
C
      call ztof(-aaa,taill)
      DO 190 I = 1,NMAX
          RCONT = 0.0
c          RCONT2 = 0.0
          SDIFF = 0.0
          SDIF2 = 0.0
          ISTP = 0
          MARK2 = IRCODE(I)
          MARK3 = IRCODE(I+1)
C
C  CASE:  K .LT. I
C         SAME ROW, DIFFERENT COLUMN
C
          K = I - 1
   30     IF (K.EQ.0) GO TO 70
              RCOL1 = C(K,I)
              RCOL2 = C(K,I+1)
C
C         RCOL1, RCOL2  -  TEMPORARY VARIABLES
C
              MARK1 = IRCODE(K)
              MARKER = 0
              MARKA = 0
              IF (IUNIQ(MARK1,2).EQ.1) MARKER = 1
              IF (IUNIQ(MARK2,2).EQ.1 .OR. IUNIQ(MARK3,2).EQ.1)
     +            MARKA = 1
              IF (MARKER.EQ.1 .AND. MARKA.EQ.1) GO TO 60
              IF (RCOL1.LT.-3.0 .OR. RCOL1.GT.3.0) GO TO 60
              IF (RCOL2.LT.-3.0 .OR. RCOL2.GT.3.0) GO TO 60
              QRCOL1 = RCOL1
              CALL ZTOF (QRCOL1,P)
              ss1=c(i,k)/(1.-p)
              dif1=c(k,i+1)-c(k,i)
              if(dif1.ne.0.0.and.c(k,i).eq.aaa) then
              ss1=c(i,k)/taill
              p=(ss1-0.5)/ss1
              call ftoz(p,rcol1)
              endif
              RFAC = 1.0
              IF (MARKER.EQ.1 .OR. IUNIQ(MARK2,2).EQ.1) RFAC = 0.5
              W1 = ss1 * EXP (-RCOL1**2) / RFAC
              W1 = W1 / (2.0 * PI * P * (1.-P))
              RCOL1 = RCOL1 * SQRT (RFAC)
              QRCOL2 = RCOL2
              CALL ZTOF (QRCOL2,P)
              ss2=c(i+1,k)/(1.-p)
              if(dif1.ne.0.0.and.c(k,i+1).eq.aaa) then
              ss2=c(i+1,k)/taill
              p=(ss2-0.5)/ss2
              call ftoz(p,rcol2)
              endif
              RFAC = 1.0
              IF (IUNIQ(MARK3,2).EQ.1 .OR. MARKER.EQ.1) RFAC = 0.5
              W2 = ss2 * EXP (-RCOL2**2) / RFAC
              W2 = W2 / (2.0 * PI * P * (1.-P))
              RCOL2 = RCOL2 * SQRT (RFAC)
              WW = (W1 * W2) / (W1 + W2)
c              RCONT = RCONT + WW
c              RDIFF = (RCOL2 - RCOL1) * WW
c              RDIF2 = ((RCOL2 - RCOL1) ** 2) * WW
              IF (c(k,i).EQ.AAA .AND. c(k,i+1).EQ.aaa) GO TO 40
                  GO TO 50
   40         ISTP = ISTP + 1
c              RCONT2 = RCONT2 + ((W1*W2) / (W1+W2))
   50         IF (c(k,i).NE.AAA .OR.c(k,i+1).NE.AAA) ISTP = 0
c              IF (RCOL1.NE.AAA .OR. RCOL2.NE.AAA) RCONT2 = 0.0
c              IF (ISTP.EQ.5) RCONT = RCONT - RCONT2
              IF (ISTP.EQ.5) GO TO 70
              rdiff=(rcol2-rcol1)*ww
              rdif2=((rcol2-rcol1)**2)*ww
              SDIFF = SDIFF + RDIFF
              SDIF2 = SDIF2 + RDIF2
              if(rdiff.eq.0.0.and.c(k,i).eq.aaa) ww=0.0
              rcont=rcont+ww
c      if(i.eq.3) write(*,*) rcol1,w1,rcol2,w2,rdiff,ww
   60         K = K - 1
              GO TO 30
C
C  CASE:  K .EQ. I
C
   70 continue
c   70     IF (ISTP.GT.0 .AND. ISTP.LT.5) RCONT = RCONT - RCONT2
          CAA = C(I,I+1)
          IF (IUNIQ(MARK2,2).EQ.1.AND.IUNIQ(MARK3,2).EQ.1)
     +    WRITE (UNIT,1040)
 1040     FORMAT (/1X, '** WARNING-  ADJOINING MARKER HORIZONS'//)
          IF (CAA.GE.-AAA .AND. CAA.LE.AAA) GO TO 80
              GO TO 90
   80     QCAA = CAA
          CALL ZTOF (QCAA,P)
          ss0=c(i+1,i)/(1.-p)
          if(caa.eq.aaa) then
          ss0=c(i+1,i)/taill
          p=(ss0-0.5)/ss0
          call ftoz(p,caa)
          endif
          W1 = (ss0 * EXP (-CAA**2)) / (2.0 * PI * P * (1.-P))
          RFAC = 1.0
          IF (IUNIQ(MARK3,2).EQ.1) RFAC = 0.5
          RCONT = RCONT + W1 / RFAC
          SDIFF = SDIFF + (W1 * CAA * SQRT (RFAC))
          SDIF2 = SDIF2 + (W1 * (CAA * RFAC)**2)
   90     ISTP = 0
c          RCONT2 = 0.0
          K = I+2
  100     IF (K.GT.MMAX) GO TO 140
C
C      CASE:  K .GT. I
C             DIFFERENT ROW, SAME COLUMN
C
              RCOL1 = C(I,K)
              RCOL2 = C(I+1,K)
              MARK1 = IRCODE(K)
              MARKER = 0
              IF (IUNIQ(MARK1,2).EQ.1) MARKER = 1
              IF (MARKER.EQ.1 .AND. MARKA.EQ.1) GO TO 130
              IF (RCOL1.LT.-3.0 .OR. RCOL1.GT.3.0) GO TO 130
              IF (RCOL2.LT.-3.0 .OR. RCOL2.GT.3.0) GO TO 130
              QRCOL1 = RCOL1
              CALL ZTOF (QRCOL1,P)
              ss1=c(k,i)/(1.-p)
              dif1=c(i+1,k)-c(i,k)
              if(dif1.ne.0.0.and.c(i,k).eq.aaa) then
              ss1=c(k,i)/taill
              p=(ss1-0.5)/ss1
              call ftoz(p,rcol1)
              endif
C
C           SET RFAC FOR "W1";  THEN AGAIN FOR "W2"
C
              RFAC = 1.0
              IF (MARKER.EQ.1 .OR. IUNIQ(MARK2,2).EQ.1) RFAC = 0.5
              W1 = ss1 * EXP (-RCOL1**2) / RFAC
              W1 = W1 / (2.0 * PI * P * (1.-P))
              RCOL1 = RCOL1 * SQRT (RFAC)
              QRCOL2 = RCOL2
              CALL ZTOF (QRCOL2,P)
              ss2=c(k,i+1)/(1.-p)
              if(dif1.ne.0.0.and.c(i+1,k).eq.aaa) then
              ss2=c(k,i+1)/taill
              p=(ss2-0.5)/ss2
              call ftoz(p,rcol2)
              endif
              RFAC = 1.0
              IF (IUNIQ(MARK3,2).EQ.1 .OR. MARKER.EQ.1) RFAC = 0.5
              W2 = ss2 * EXP (-RCOL2**2) / RFAC
              W2 = W2 / (2.0 * PI * P * (1.-P))
              RCOL2 = RCOL2 * SQRT (RFAC)
              WW = (W1 * W2) / (W1 + W2)
c              RCONT = RCONT+WW
c              RDIFF = (RCOL1 - RCOL2) * WW
c              RDIF2 = (RCOL1 - RCOL2)**2 * WW
              IF (c(i,k).EQ.AAA .AND. c(i+1,k).EQ.aaa) GO TO 110
                  GO TO 120
  110         ISTP = ISTP + 1
c              RCONT2 = RCONT2 + ((W1*W2) / (W1+W2))
  120         IF (c(i,k).NE.AAA .OR.c(i+1,k).NE.AAA) ISTP = 0
c              IF (RCOL1.NE.AAA .OR. RCOL2.NE.AAA) RCONT2 = 0.0
c              IF (ISTP.EQ.5) RCONT = RCONT - RCONT2
              IF (ISTP.EQ.5) GO TO 140
              rdiff=(rcol1-rcol2)*ww
              rdif2=((rcol1-rcol2)**2)*ww
              SDIFF = SDIFF + RDIFF
              SDIF2 = SDIF2 + RDIF2
              if(rdiff.eq.0.0.and.c(i,k).eq.aaa) ww=0.0
              rcont=rcont+ww
  130         K = K + 1
              GO TO 100
  140         continue
c  140     IF (ISTP.GT.0 .AND. ISTP.LT.5) RCONT = RCONT - RCONT2
C
C         SET QDIFF, VARX, STDDEV, RDENO
C
          QDIFF = 0.0
          IF (RCONT.NE.0.0)  QDIFF = SDIFF / RCONT
          RDENO = (MPAIR(I) - 1) * RCONT
          VARX = 0.0
          IF (RDENO.NE.0.0)  VARX = (-QDIFF*QDIFF*RCONT + SDIF2) / RDENO
          STDDEV = CONST3
          IF (VARX.GE.0.0)  STDDEV = SQRT (VARX)
          IF (RDENO.EQ.0.0 .OR. MPAIR(I).EQ.1)  STDDEV = CONST4
C
          IF (INEG.NE.1)  GO TO 182
              IF (STDDEV.EQ.CONST4 .OR. STDDEV.EQ.CONST3)  QDIFF = 0.0
c  Test added on June 20, 1995
c  crit2 is replaced by crit3; weight should exceed 3.5
              crit3=10.0/crit2
              ik=0
              IF (FLOAT (MPAIR(I)).LT.CRIT3)  then
              ik=1
              QDIFF = 0.0
              endif
              if(rcont.lt.3.5) qdiff=0.0
  182     QDAR(I+1) = QDAR(I) + QDIFF
          st(i)=stddev
          IF (LLL.NE.0)  WRITE (UNIT,2000) I, IRCODE(I), IRCODE(I+1),
     +     QDIFF, QDAR(I+1), SDIFF, mpair(I), RCONT, STDDEV
 2000     FORMAT (1X, I4, I5, '-', I3, F11.4,F12.4,F12.4,i7,F7.1,F10.4)
  190 CONTINUE
      RETURN
      END
      SUBROUTINE XUNIQ1 (I,IX1)
C
C ... SUBROUTINE XUNIQ1 IS USED WITH SUBROUTINE HPFILT TO CONSTRUCT
C  A MATRIX OF POSITION REFERENCES FOR ALL SELECTED UNIQUE EVENTS.
C  THIS MATRIX IS USED BY SUBROUTINE XUNIQ2 TO PLACE UNIQUE
C  EVENTS INTO THE FINAL SEQUENCE.
C
C  ACCEPTS:  I     - THE WELL NUMBER FOR "NUNIQ"
C            IX1   - THE ORIGINAL (UNFILTERED) SEQUENCE DATA
C  RETURNS:  MUNIQ - ROWS OF ORDERED PAIRS (SEE BELOW)
C
C  HERE, "IX2" SHOULD CONTAIN THE FILTERED SEQUENCE DATA
C            (FILTERED BUT NOT RECODED)
C
C  CALLED BY SUBROUTINE HPFILT.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS, MMAX, IX2(KDIM1,KDIM3), MM(KDIM4), M(KDIM4),
     +        RMAT(KDIM6,KDIM6)
      COMMON /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      INTEGER IX1(KDIM1,KDIM3), IFIL(KDIM4), NOTFIL(KDIM4)
C
C  MUNIQ(KDIM1, 2 * MAX.NR.OF UNIQUE EVENTS) -
C      ROW NUMBER REFLECTS WELL NUMBER.
C      ORDERED PAIR:  FIRST ELEMENT IS ORIGINAL CODE NR. OF UQ. EVT.
C                     SECOND ELEMENT IS LOWEST NUMBERED NEIGHBOUR
C                       AT THE SAME LEVEL OR A NEARBY LEVEL.
C                       (SEE BODY OF ROUTINE FOR DETAILS)
C
C  IFIL(KDIM4)   - VECTOR OF LEVEL NUMBERS FOR THE FILTERED SEQUENCE
C  NOTFIL(KDIM4) - VECTOR OF LEVEL NUMBERS FOR THE UNFILTERED SEQ.
C
      DO 10 J = 1,N
          IFIL(J) = 0
          NOTFIL(J) = 0
   10 CONTINUE
C
C     ASSIGNMENT OF LEVEL NUMBERS, STARTING AT "1" AND INCREASING
C         BY 1, PANNING LEFT-TO-RIGHT IN THE GIVEN SEQUENCE.
C
C     KK = LEVEL NUMBER
      KK = 0
      DO 30 J = 1,KDIM3
          ID = IX1(I,J)
          IF (ID.EQ.0) GO TO 40
          IF (ID.LT.0) GO TO 20
          KK = KK + 1
   20     NOTFIL(IABS (ID)) = KK
   30 CONTINUE
   40 MAXTET = KK
      KK = 0
      DO 60 J = 1,KDIM3
          ID = IX2(I,J)
          IF (ID.EQ.0) GO TO 70
          IF (ID.LT.0) GO TO 50
          KK = KK + 1
   50     IFIL(IABS (ID)) = KK
   60 CONTINUE
   70 CONTINUE
      KKK = 1
      DO 120 J = 1,KDIM3
C
C         LOCATE NEXT UNIQUE EVENT;  SET "LEV1"
C
          ID = IX1(I,J)
          IF (ID.EQ.0) GO TO 900
          IDA = IABS (ID)
          IF (IUNIQ(IDA,1).NE.1) GO TO 120
          LEV1 = NOTFIL(IDA)
C
          DO 80 JJ = 1,N
C
C             LOOK IN THE FILTERED SEQUENCE FOR A (NON-UNIQUE) EVENT
C             ON THE  *SAME*  LEVEL.  (TAKE THE LOWEST NUMBERED FOSSIL)
C
              IF (NOTFIL(JJ).NE.LEV1) GO TO 80
              IF (IFIL(JJ).EQ.0) GO TO 80
              MUNIQ(I,KKK) = IDA
              KKK = KKK + 1
              MUNIQ(I,KKK) = -JJ
              KKK = KKK + 1
              GO TO 120
   80     CONTINUE
          ISTEP = LEV1
          ISTEM = LEV1
C
C    ELSE LOOK FOR LOWEST NUMBERED FOSSIL AT AN  *ADJACENT*  LEVEL.
C         FIRST TRY 1-BELOW, 1-ABOVE;
C         THEN (IF NECESSARY)  2-BELOW, 2-ABOVE;   ... AND SO ON.
C
   90     ISTEP = ISTEP + 1
          ISTEM = ISTEM - 1
          DO 110 JJ = 1,N
              IDSTP = NOTFIL(JJ)
              IF (ISTEM.LE.0 .OR. IDSTP.NE.ISTEM) GO TO 100
                  IF (IFIL(JJ).EQ.0) GO TO 100
                  MUNIQ(I,KKK) = -IDA
                  KKK = KKK + 1
                  MUNIQ(I,KKK) = JJ
                  KKK = KKK + 1
                  GO TO 120
  100         IF (ISTEP.GT.MAXTET .OR. ISTEP.NE.IDSTP) GO TO 110
                  IF (IFIL(JJ).EQ.0) GO TO 110
                  MUNIQ(I,KKK) = IDA
                  KKK = KKK + 1
                  MUNIQ(I,KKK) = JJ
                  KKK = KKK + 1
                  GO TO 120
  110     CONTINUE
          IF (ISTEP.LE.MAXTET .OR. ISTEM.GT.0) GO TO 90
  120 CONTINUE
  900 RETURN
      END
      SUBROUTINE XUNIQ2 (N,II,ICNT,IVEC,RMAT,RUNIQ,kuniq)
C
C ... SUBROUTINE TO PLACE UNIQUE EVENTS IN FINAL SEQUENCE
C
C  POSITIONS OF UNIQUE EVENTS ARE DETERMINED FROM FIRST AND SECOND
C  APPROXIMATIONS BASED ON POSITIONS OF NEIGHBOURING EVENTS.
C
C  ACCEPTS:  II    - THE WELL NUMBER
C            IVEC(KDIM3)  - THE SEQUENCE DATA FOR THIS WELL,
C                  RE-ORDERED ACCORDING TO THE OPTIMUM SEQUENCE
C            ICNT  = (LENGTH OF THIS SEQUENCE) + 1
C            RMAT(KDIM3, 3)  - COL. 1:  CUMULATIVE DISTANCES;
C                              COL. 2:  FIRST ORDER DIFFERENCES;
C                              COL. 3:  SECOND ORDER DIFFERENCES
C                                 SLIGHTLY TRANSFORMED IN COMP().
C  RETURNS:  RUNIQ  -  COLUMN 1 SUMS UP ALL 'X2'S
C                      COLUMN 2 COUNTS HOW MANY 'X2'S THERE WERE
C
C  CALLED IN MAIN AND BY SUBROUTINE COMP.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      INTEGER ILEV(KDIM4), IVEC(KDIM4)
      REAL    RMAT(KDIM4,3), RUNIQ(KDIM4,2), RX(3,2)
      DATA    CONST5/1.386/
C
C  RX(3,2) -   THREE ORDERED PAIRS
C      1:  (SUM OF CUM. DISTANCES;  COUNT OF CUM. DISTANCES)
C      2:  (SUM OF 1ST-ORD. DIFFERENCES;  COUNT OF 1-O.D.)
C      3:  (SUM OF 2ND-ORD. DIFFERENCES;  COUNT OF 2-O.D.)
C
      DO 10 I = 1,N
          ILEV(I) = 0
   10 CONTINUE
      KK = 0
      DO 30 I = 1,ICNT
          IF (IVEC(I).LT.0) GO TO 20
          KK = KK + 1
   20     ILEV(I) = KK
   30 CONTINUE
      MAXTEP = KK
      ITEM = 1
C
C  LOOP UNTIL NO MORE UNIQUE EVENTS
C        (MAIN LOOP)
C
   40 IEVENT = MUNIQ(II,ITEM)
      IREFF = MUNIQ(II,ITEM+1)
      IF (IEVENT.EQ.0) GO TO 999
      DO 50 I = 1,N
          IF (IABS(IVEC(I)) .EQ. IABS(IREFF)) GO TO 60
   50 CONTINUE
   60 LEVEL = ILEV(I)
      IF (IREFF.GE.0) GO TO 70
          LEV0 = LEVEL
          LEVM = LEV0 - 1
          LEVP = LEV0 + 1
          GO TO 90
   70 IF (IEVENT.GE.0) GO TO 80
          LEVM = LEVEL
          LEVP = LEVEL + 1
          LEV0 = 0
          GO TO 90
   80 LEV0 = 0
          LEVP = LEVEL
          LEVM = LEVEL - 1
C
C  CALCULATE THE FIRST APPROXIMATION, X1
C
   90 DO 100 I = 1,3
          RX(I,1) = 0.0
          RX(I,2) = 0.0
  100 CONTINUE
      DO 130 I = 1,ICNT
          LEVI = ILEV(I)
          IF (LEVM.LE.0) GO TO 110
              IF (LEVM.NE.LEVI) GO TO 110
              RX(1,1) = RX(1,1) + RMAT(I,1)
              RX(1,2) = RX(1,2) + 1.0
              GO TO 130
  110     IF (LEV0.EQ.0) GO TO 120
              IF (LEV0.NE.LEVI) GO TO 120
              RX(2,1) = RX(2,1) + RMAT(I,1)
              RX(2,2) = RX(2,2) + 1.0
              GO TO 130
  120     IF (LEVP.GT.MAXTEP) GO TO 130
              IF (LEVP.NE.LEVI) GO TO 130
              RX(3,1) = RX(3,1) + RMAT(I,1)
              RX(3,2) = RX(3,2) + 1.0
  130 CONTINUE
      RRR = 0.0
      ARX1 = 0.0
      ARX2 = 0.0
      ARX3 = 0.0
C
C     COMPUTE AVERAGES
C
      IF (RX(1,2).EQ.0.0) GO TO 140
          ARX1 = RX(1,1) / RX(1,2)
          RRR = RRR + 1.0
  140 IF (RX(2,2).EQ.0.0) GO TO 150
          ARX2 = RX(2,1) / RX(2,2)
          RRR = RRR + 1.0
  150 IF (RX(3,2).EQ.0.0) GO TO 160
          ARX3 = RX(3,1) / RX(3,2)
          RRR = RRR + 1.0
  160 X1 = (ARX1 + ARX2 + ARX3) / RRR
C
C  CALCULATE THE SECOND APPROXIMATION,  X2
C
      RANGE1 = X1 - CONST5
      RANGE2 = X1 + CONST5
c Section inserted on 21 June, 1995
      if(kuniq.eq.1) then
         range1=x1-3.0
         range2=x1+3.0
      endif
c End of insert
      RSUM = 0.0
      RRR = 0.0
      DO 190 I = 1,ICNT
          DIST1 = RMAT(I,1)
          IF (DIST1.LT.RANGE1 .OR. DIST1.GT.RANGE2) GO TO 190
          IF (LEV0.EQ.0 .OR. LEVEL.NE.ILEV(I)) GO TO 170
              RSUM = RSUM + DIST1
              RRR = RRR + 1
              GO TO 190
  170     IF (ILEV(I).LE.LEVM .AND. LEVM.GT.0) GO TO 180
              IF (ILEV(I).LT.LEVP .OR. LEVP.GT.MAXTEP) GO TO 190
              RSUM = RSUM + ((DIST1 + RANGE1) * 0.5)
              RRR = RRR + 1
              GO TO 190
  180     RSUM = RSUM + ((DIST1 + RANGE2) * 0.5)
              RRR = RRR + 1
  190 CONTINUE
      X2 = X1
      IF (RRR.NE.0.0)  X2 = RSUM / RRR
      ID = IABS (IEVENT)
      RUNIQ(ID,1) = RUNIQ(ID,1) + X2
      RUNIQ(ID,2) = RUNIQ(ID,2) + 1.0
      ITEM = ITEM + 2
      GO TO 40
C
  999 RETURN
      END
      SUBROUTINE FPAGE
C
      WRITE(61,'(4(/))')
      WRITE(61,'(6X,A)')         
     +'         RRRR      A      SSS     CCC         PPPP     CCC'       
      WRITE(61,'(6X,A)')
     +'         R   R    A A    S   S   C   C        P   P   C   C'      
      WRITE(61,'(6X,A)') 
     +'         R   R   A   A   S       C            P   P   C'          
      WRITE(61,'(6X,A)')  
     +'         RRRR    AAAAA    SSS    C      ===== PPPP    C'          
      WRITE(61,'(6X,A)')    
     +'         R R     A   A       S   C            P       C'          
      WRITE(61,'(6X,A)')  
     +'         R  R    A   A   S   S   C   C        P       C   C'      
      WRITE(61,'(6X,A)')    
     +'         R   R   A   A    SSS     CCC         P        CCC'
      WRITE(61,'(2(/))')
      WRITE(61,'(6X,A)')
     +'                 RASC Version 18 (2002)'
      write(61,'(1(/))')
      write(61,'(6x,a)')
     +'                            by'
      write(61,'(1(/))')
      write(61,'(6x,a)')
     +'                 F.P. Agterberg and F.M. Gradstein'
      write(61,'(1(/))')
      write(61,'(6x,a)')
     +'               RANKING AND SCALING OF FOSSIL EVENTS'
      write(61,'(6x,a)')
     +'               ____________________________________'
      WRITE(61,'(5(/))')       
      WRITE(61,'(6X,A)')        
     +'RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR'//
     +'RRRRRRRRRRRRRRRRRRR'    
      WRITE(61,'(/)')        
      WRITE(61,'(6X,A)')    
     +'                            REFERENCES'
      WRITE(61,'(6X,A)') 
     +' '
      WRITE(61,'(6X,A)')    
     +'   F.P.AGTERBERG & L.D.NEL, 1982A. ALGORITHMS FOR THE RANKING'
      write(61,'(6x,a)')
     +'   OF STRATIGRAPHIC EVENTS, COMPUTERS & GEOSCIENCES,'
      WRITE(61,'(6X,A)')    
     +'   VOL. 8, NO. 1, P. 69-90.'
      write(61,'(/6x,a)')
     +'   F.P.AGTERBERG & L.D.NEL, 1982B. ALGORITHMS FOR THE SCALING'
      write(61,'(6x,a)')
     +'   OF STRATIGRAPHIC EVENTS, COMPUTERS & GEOSCIENCES,'
      write(61,'(6x,a)')
     +'   VOL. 8, NO. 2, P. 163-189.'
      WRITE(61,'(/6X,A)')
     +'   F.M.GRADSTEIN & F.P.AGTERBERG, 1982. MODELS OF CENOZOIC'
      write(61,'(6x,a)')
     +'   FORAMINIFERAL STRATIGRAPHY, NORTWESTERN ATLANTIC MARGIN.'
      WRITE(61,'(6X,A)')    
     +'   IN: QUANTITATIVE STRATIGRAPHIC CORRELATION, EDITORS:'
      WRITE(61,'(6X,A)')    
     +'   J.M.CUBITT & R.A.REYMENT, WILEY, CHICHESTER, P. 119-173.'
      WRITE(61,'(/6X,A)')
     +'   F.M.GRADSTEIN, F.P.AGTERBERG, J.C.BROWER &'
      WRITE(61,'(6X,A)')    
     +'   W.SCHWARZACHER, 1985. QUANTITATIVE STRATIGRAPHY,'
      WRITE(61,'(6X,A)')    
     +'   D. REIDEL PUBL. CO. & UNESCO, 598 PP.'
      WRITE(61,'(/6X,A)')
     +'   F.P.AGTERBERG, 1990. AUTOMATED STRATIGRAPHIC CORRELATION,'
c     +'   M.HELLER, W.S.GRADSTEIN, F.M.GRADSTEIN &'//
c     +' F.P.AGTERBERG & S.N.LEW'
      WRITE(61,'(6X,A)')
     +'   ELSEVIER, AMSTERDAM, 424 PP.'
c     +'         1985. GEOL. SURVEY CANADA REPT. 1203, 62 PP. (USER`S'
      WRITE(61,'(/6X,A)')
     +'   F.M.GRADSTEIN & F.P.AGTERBERG, IN PRESS. UNCERTAINTY IN'
      write(61,'(6x,a)')
     +'   STRATIGRAPHIC CORRELATION. IN PROCEEDINGS, HIGH RESOLUTION'
      write(61,'(6x,a)')
     +'   SEQUENCE STRATIGRAPHY CONFERENCE OF NORWEGIAN PETROLEUM'
      write(61,'(6x,a)')
     +'   SOCIETY, STAVANGER, NOVEMBER 1995, ELSEVIER, AMSTERDAM.'
c     +'          MANUAL OF ORIGINAL VERSION OF RASC).'
      WRITE(61,'(/)')
      WRITE(61,'(6X,A)')        
     +'RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR'//
     +'RRRRRRRRRRRRRRRRRRR'    


      WRITE(61,'(6(/))')        
      WRITE(61,100)
100   FORMAT(1X,T6,    
     + '  EXPLANATION OF RASC17 PARAMETERS FILE (*.INP)'/
     +T6, '  ---------------------------------------------'/
     + ' '/     
     + ' '/      
     +T6, '  RECORD NO. 1     RUN PARAMETERS   '/
     +T6, '  ------------     -------------- '/
     + ' ')
      WRITE(61,120)
120   FORMAT(1X,T8,    
     +'NS     -  (INTEGER)  NUMBER OR SEQUENCES OR WELLS (NS  <=100)'/
     +T6, '  IOCR   -  (INTEGER)  ELEMENTS OCCURRING FEWER THAN "IOCR"'/     
     +T8, '                     TIMES IN THE DATA SET WILL BE IGNORED'/
     +T6, '  INIQ   -  (INTEGER)  = 1, IF UNIQUE EVENTS OR MARKER'/ 
     +T10, '                   HORIZONS ARE INCLUDED (RECORDS 2 AND 3)'/
     +T6, '  IALPHA   = 0,        TERMINATION AFTER RANKING SOLUTION'/
     +T6, '           = 1,        SCALING ANALYSIS WILL BE DONE'/
     +T6, '  ICASC    = 0,        NO WELL DATA OUTPUT FILE'/
     +T6, '           = 1,        OUTPUT FILE FOR USE AS INPUT TO CASC'/
     +T6, '                       PROGRAM'/
c     +T8, 'ITER   -  (INTEGER)  MAXIMUM NUMBER OF CUMULATIVE ORDER'/
c     +T6, '                       MATRIX TRANSFORMATIONS'/
c     +T14, '               ALLOWED  (EXAMPLE: 12000)'/
c     +T8, 'CRIT1  -  (REAL)     TRANSPOSE ELEMENTS WITH SUM LESS THAN'/
c     +T12,'                 "CRIT1" IN THE ORDER MATRIX WILL BE ZEROED'/
c
c     +T14, '               BEFORE THE RANKING SOLUTION.(CRIT1 <= IOCR)'/
c     +T6, '  TOL    -  (REAL)     TOLERANCE:  S(I,J) MAY BE LOWER THAN'/
c     +T6, '                       S(J,I) BY AS MUCH AS "TOL"'/
c     +T8,'AAA    -  (REAL)     FRACTILE FOR TRUNCATION POINT OF NORMAL'/
c     +T10, '                   DISTRIBUTION.  (EXAMPLE: AAA = 1.645)'/
     +T7,' ICRIT  -  (INTEGER)  TRANSPOSE ELEMENTS WITH SUM LESS THAN')
      WRITE(61,140)
140   FORMAT(1X,    
     +T14, '               "ICRIT" IN THE ORDER MATRIX WILL BE ZEROED'/
     +T6, '                       BEFORE THE SCALING ANALYSIS.'/
c     +T6, '                       (CRIT2 >= CRIT1)'///
c     +T6 ,'  RECORD NO. 2     PROCESSING CONTROL      (FIXED FORMAT)'/
c     +T6, '  ------------     ------------------'/
c     +T6, ' '/
c     +T6, '  ALL 12 PARAMETERS ON THIS RECORD ARE INTEGERS. FOR'/
c     +T8, 'SHORT VERSION OF RASC (RANKING ALGORITHMS ONLY),IALPHA=0')
c      WRITE(61,'(2(/))')
c      WRITE(61,160)
c160   FORMAT(1X,T6,
c     +T6, '  ITAPE    = 1   FOR DATA TO BE READ FROM "TAPE10"'/
c     +T6, '            ELSE,       DATA WILL BE READ FROM RECORDS'/
c     +T6, '                 IMMEDIATELY FOLLOWING RECORD NO. 4'/
c     +T8, 'IOMAT    = 1   FOR PRINTOUT OF ORDER AND FREQUENCY MATRICES'/
c
c     +T6, '                 AS WELL AS INTERMEDIATE TABLES;'/
c     +T6, '            ELSE,       THESE OUTPUTS WILL BE SUPPRESSED'/
c     +T6, '  ISRT     = 0,  MODIFIED HAY METHOD WITHOUT PRESORTING'/
c     +T6, '             1,  DATA WILL BE PRE-SEQUENCED FOR OPTIMIZED'/
c     +T6, '                 STARTING SEQUENCE'/
c     +T6, '            ELSE,       CONDENSED OPTIMUM SEQUENCE AFTER'/
c     +T6, '                 PRESORTING'/
c     +T14, '    ELSE,       TERMINATION AFTER RANKING SOLUTION, BUT'/
c     +T23, 'STEPWISE SEQUENCING PROGRESS WILL BE'/
c     +T6, '                 PRINTED BEFORE TERMINATION.'/
c     +T7, ' ITAB1    = 1,  AN OCCURRENCE TABLE FOR THE WELLS IS TO BE'/
c
c     +T6, '                 PRINTED'/
c     +T6, '            ELSE,       NO TABLE.'/
c     +T8, 'ISCORE   = 1,  STEP MODEL COMPARISON OF INDIVIDUAL WELLS')
c      WRITE(61,180)
c180   FORMAT(1X,
c     +T12, '           AND FOSSILS WITH OPTIMUM SEQUENCE IS PERFORMED'/
c     +T6, '  ICOMP    = 1   FOR NORMALITY TESTS ON INDIVIDUAL WELLS'/
c     +T6, '  ISKIP    = 1   IF CUMULATIVE ORDER MATRIX IS TO BE USED'/
c     +T10, '             (RANKING SOLUTION WILL BE BASED ON PRESORTING'/
c     +T6, '                 ONLY)'/
c     +T8, '          ELSE,       RASC WILL GO AHEAD AND PERFORM MATRIX'/
c     +T6, '                 PERMUTATIONS.'/
c     +T6, '  IFIN     = 1   FOR APPLICATION OF FINAL RE-ORDERING'/
c     +T6, '  INOSC    = 0,  NO SCALING OUTPUT;'/
c     +T6, '           = 1,  WEIGHTED DISTANCE OUTPUT ONLY;'/
c     +T6, '            ELSE,       WEIGHTED AND UNWEIGHTED OUTPUT.'/
c     +T6, '  INEG     = 1,  LARGE DISTANCES FOR SMALL SAMPLES WILL BE'/
c     +T6, '                 SUPPRESSED;'/
c     +T6, '            ELSE,      NO SUPPRESSION.'/
c     +T6, '  ISCAT    = 1,  SCATTERGRAMS;'/
c     +T6, '            ELSE,      NO SCATTERGRAMS.'/
c     +T6, '  IVAR     = 1,  VARIANCE ANALYSIS TO BE PERPORMED FOR EACH'/
c     +T6, '                 WELL;'/
c     +T6, '            ELSE,      NO VARIANCE ANALYSIS.'/
c     +T6, '            ELSE,      WELL DATA OUTPUT FILE.'/
c     + ' '/
c     +T6, '  RECORD NO. 3    OUTPUT DIRECTION (FIXED FORMAT)'/
c     +T6, '  ------------    ------------------------------'/
c     + ' ')
c      WRITE(61,'(6X,A)')
c     + '  THIS RECORD IS MADE UP OF 11 INTEGER VALUES (11I2 FORMAT)'
c      WRITE(61,'(6X,A)')
c     + '  WHICH DIRECT OUTPUT FROM THE 11 DIFFERENT PROGRAM SECTIONS.'
c      WRITE(61,'(6X,A)')
c     + '  THESE VALUES ARE STORED IN THE PROGRAM IN THE ARRAY OUT().'
c      WRITE(61,'(6X,A)')
c     + '  THE FIRST VALUE REFERS TO SECTION 1, THE SECOND TO SECTION'
c      WRITE(61,'(6X,A)')
c     + '  2 AND SO ON. IF THE VALUE IS SET TO 1, THE CORRESPONDING'
c      WRITE(61,'(6X,A)')
c     + '  SECTION`S OUTPUT WILL BE DIRECTED TO THE MAIN '//
c     + 'OUTPUT FILE.'
c      WRITE(61,'(6X,A)')
c     + '  OTHERWISE THE SECTIONS OUTPUT WILL GO TO THE'//
c     +' EXTRA OUTPUT FILE.'
c      WRITE(61,300)
c300   FORMAT(1X,/
c     + T6,'  OUT(1) - TABULATION OF EVENT OCCURRENCES',
c     +    ' (DISABLED IN RASC15)'/
c     + T6,'  OUT(2) - MODIFIED SEQUENCE DATA',
c     +    ' (DISABLED IN RASC15)'/
c     + T6,'  OUT(3) - DICTIONARIES'/
c     + T6,'  OUT(4) - CYCLES'/
c     + T6,'  OUT(5) - OPTIMUM SEQUENCE'/
c     + T6,'  OUT(6) - OCCURRENCE TABLE & STEP MODEL'/
c     + T6,'  OUT(7) - SCATTERGRAMS USING OPTIMUM SEQUENCE'/
c     + T6,'  OUT(8) - SCALING (WEIGHTED)'/
c     + T6,'  OUT(9) - SCALING AFTER 5 ITERATIONS'/
c     + T6,'  OUT(10)- NORMALITY TEST'/
c     + T6,'  OUT(11)- UNIQUE (RARE) EVENTS IN OPTIMUM SEQUENCE'/
     + T6,' '/
     + T6,'   RECORD NO. 2    UNIQUE (RARE) EVENTS'/
     + T6,'   ------------    --------------------'/
     + ' ')  
      WRITE(61,'(6X,A)')    
     + '  IF INIQ = 1, UP TO 20 UNIQUE EVENTS WILL BE READ FROM THIS'
      WRITE(61,'(6X,A)')    
     + '  RECORD IN 20I4 FORMAT.  IF NO UNIQUE EVENTS ARE REQUESTED,' 
      WRITE(61,'(6X,A)')    
     + '  THIS RECORD IS LEFT BLANK.'
c      WRITE(61,'(6X,A)')
c     + '  "ICOMP" MUST EQUAL 1.'
      WRITE(61,'(6X,A)')    
     + ' '
      WRITE(61,'(6X,A)')        
     + '  RECORD NO. 3     MARKER HORIZONS'
      WRITE(61,'(6X,A)')    
     + '  ------------     ---------------'
      WRITE(61,'(6X,A)')    
     + ' ' 
      WRITE(61,'(6X,A)')    
     + '  IF INIQ = 1, UP TO 20 MARKER HORIZONS WILL BE READ FROM'
      WRITE(61,'(6X,A)')    
     + '  THIS RECORD IN 20I4 FORMAT.  IF NO MARKER HORIZONS ARE' 
      WRITE(61,'(6X,A)')    
     + '  REQUESTED, THIS RECORD IS LEFT BLANK.'
      WRITE(61,'(2(/))')    
      WRITE(61,'(6X,A)')    
     + '  EXPLANATION OF INPUT DATA FORMAT (FIXED FORMAT)' 
      WRITE(61,'(6X,A)')    
     + '  ----------------------------------------------'  
      WRITE(61,'(6X,A)')    
     + ' '  
      WRITE(61,'(6X,A)')    
     + '  OBSERVED SEQUENCES OF EVENTS:'  
      WRITE(61,'(6X,A)')    
     + '    THE OBSERVED DATA ARE GIVEN AS "NS" SEQUENCES.  '//
     + 'A "SEQUENCE"'  
      WRITE(61,'(6X,A)')    
     + '    CONSISTS OF A TITLE RECORD (IN 5A4 FORMAT) FOLLOWED BY ONE'
      WRITE(61,'(6X,A)')    
     + '    OR  MORE RECORDS OF SEQUENCE DATA (IN MULTIPLE I4 FORMAT).' 
      WRITE(61,'(6X,A)')    
     + '    EACH OF THESE DATA RECORDS IS PARTITIONED INTO 20 FIELDS,'
      WRITE(61,'(6X,A)')    
     + '    EACH FIELD HAVING I4 FORMAT. THUS, THERE MAY BE UP TO 20' 
      WRITE(61,'(6X,A)')    
     + '    EVENTS IN ONE RECORD. RASC17 ALLOWS UP TO 100 '//
     + 'SEQUENCES (WELLS)'
      WRITE(61,'(6X,A)')       
     + '    AND UP TO 300 EVENTS PER SEQUENCE'    
      WRITE(61,'(6X,A)')    
     + ' ' 
      WRITE(61,'(6X,A)')    
     + '  NOTE:'  
      WRITE(61,'(6X,A)')    
     + '    A SEQUENCE IS DEEMED TO HAVE ENDED WHEN A FIELD' 
      WRITE(61,'(6X,A)')    
     + '    CONTAINING   -999   IS ENCOUNTERED.'
c      WRITE(61,'(6X,A)')
c     + '    THE SEQUENCES ARE READ EITHER FROM "TAPE10" (IF ITAPE = 1)'
c      WRITE(61,'(6X,A)')
c     + '    OR FROM RECORDS IMMEDIATELY FOLLOWING RECORD NUMBER 4  (IF'
c      WRITE(61,'(6X,A)')
c     + '    ITAPE = 0).'
      WRITE(61,'(6X,A)')
     + ' ' 
      WRITE(61,'(6X,A)')    
     + '  DICTIONARY:'   
      WRITE(61,'(6X,A)')    
     + '    THE DICTIONARY IS MERELY A LIST OF EVENT LABELS:  ONE LABEL'  
      WRITE(61,'(6X,A)')    
     + '    PER RECORD, 10A4 FORMAT.  AN EXTRA RECORD WITH THE LABEL'
      WRITE(61,'(6X,A)')    
     + '    `LAST` IN THE FIRST 4 COLUMNS MUST BE PLACED AT THE END.' 
      WRITE(61,'(6X,A)')    
     + '    THE LIMITATION IN RASC17 IS 998 RECORDS'//
     + ' (NOT COUNTING `LAST`)'  
      WRITE(61,'(6X,A)')    
     + '    THE DICTIONARY IS ALWAYS READ FROM "TAPE99".' 
      WRITE(61,'(6X,A)')    
     + ' '  
      WRITE(61,'(/)')         
      WRITE(61,'(6X,A)')
     + 'RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR'//
     + 'RRRRRRRRRRRR'        
    
      WRITE (61,1000)
 1000 FORMAT (//////10X, 'RESULTS OBTAINED BY MEANS OF PROGRAM ',
     +'RASC17 (1998)'/
     +10X,'FOR RANKING AND SCALING OF FOSSIL EVENTS')

      RETURN
      END

           SUBROUTINE ASORT (NC,IPOS,FOSSIL)
C
C ... PRINTS ALPHABETIC LISTING OF DICTIONARY.  USES FUNCTION LLE(),
C  "LEXICALLY LESS THAN OR EQUAL TO".  THIS ROUTINE IS SOMEWHAT
C  TIME CONSUMING AND MAY BE OMITTED BY REMOVING ITS CALL.
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS, MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      DIMENSION    IPOS(KDIM4)
      CHARACTER*4  NAME, ITITLE
      CHARACTER*40 FOSSIL(KDIM4), TEMP
C
      DO 2000 I = 1,N
          IPOS(I) = I
          WRITE (FOSSIL(I),1800) (ITITLE(I,J), J = 1,NC)
 1800     FORMAT (20A4)
 2000 CONTINUE
      JMAX = N - 1
 3000 IFLAG = 0
      DO 6000 J = 1,JMAX
          IF (LLE (FOSSIL(J), FOSSIL(J+1)) )  GO TO 6000
               ITEMP = IPOS(J)
               IPOS(J) = IPOS(J+1)
               IPOS(J+1) = ITEMP
               TEMP = FOSSIL(J)
               FOSSIL(J) = FOSSIL(J+1)
               FOSSIL(J+1) = TEMP
               IFLAG = 1
 6000 CONTINUE
      IF (IFLAG.EQ.0)  GO TO 7000
          JMAX = JMAX - 1
          GO TO 3000
C 
 7000 RETURN
      END

      SUBROUTINE CHECK (NS, CRIT1, CRIT2, IOCR, IER)
C
C ... SUBROUTINE TO CHECK THE VALUES OF THE INPUT VARIABLES.
C  ISSUES ERROR MESSAGES WHERE APPROPRIATE AND ADVISES CALLING
C  PROGRAMME TO TERMINATE EXECUTION.
C  RETURNS  "IER"
C
C  CALLED BY SUBROUTINE READIN
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      LOGICAL FLAG
C
      FLAG = .FALSE.
      IER = 0
 1100 IF (NS.GE.1 .AND. NS.LE.KDIM1)  GO TO 2200
           FLAG = .TRUE.
           IF (NS.GT.KDIM1)  GO TO 1310
                WRITE (6,1300)
 1300           FORMAT (///1X, '*** ERROR:   NS  LESS THAN 1 (ONE)')
                GO TO 2200
 1310      WRITE (6,1315) KDIM1, KDIM1
 1315           FORMAT (///1X, '*** ERROR:   NS  GREATER THAN', I4,
     +          /16X, 'FOR MORE THAN', I4, ' WELLS,')
                WRITE (6,1320)
 1320           FORMAT (11X, 'PLEASE USE A LARGER VERSION OF RASC')
C
 2200 IF (CRIT2.GE.CRIT1)  GO TO 2400
           WRITE (6,2230)
 2230      FORMAT (///1X, '** WARNING:  CRIT2 LESS THAN CRIT1',
     +     /16X, 'CRIT2 HAS BEEN SET EQUAL TO CRIT1')
           CRIT2 = CRIT1
C
 2400 IF (CRIT1 .LE. FLOAT(IOCR))  GO TO 9000
           FLAG = .TRUE.
           WRITE (6,2430)
 2430      FORMAT (///1X, '*** ERROR:   CRIT1 GREATER THAN "IOCR"')
C
 9000 IF (.NOT.FLAG)  GO TO 9900
           WRITE (6,9200)
 9200      FORMAT (///1X, 'AN ERROR HAS BEEN DETECTED IN THE INPUT',
     +     ' PARAMETERS.'//6X, '**** EXECUTION WILL BE TERMINATED.')
           IER = 1
C
 9900 RETURN
      END
      SUBROUTINE COMP (QDAR,INIQ,UNIT,kuniq,itnt)
C
C ... SUBROUTINE TO PERFORM NORMALITY TEST ON INDIVIDUAL WELLS
C  AND TO FIT UNIQUE EVENTS INTO OPTIMUM SEQUENCE
C
C  RETURNS:  MMAX   - THE REVISED MMAX FOR ORDER() AND DENDRO()
C            QDAR   - SAME AS ORIGINAL QDAR FROM WDIST(),
C                       BUT WITH UNIQUE EVENTS APPENDED.
C                       SENT TO ORDER()
C
C  OBTAINS "RUNIQ" FROM CALL TO XUNIQ2()
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      COMMON /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER IVEC(KDIM3), UNIT
      REAL    CLAS(10), CLASS(10), QDAR(KDIM6+MAXUQ), RMAT(KDIM3,3),
     +        RUNIQ(KDIM4,2), SECDIF(3000)
      CHARACTER*4  NAME, ITITLE, IBLANK, IASK1, IASK2, ISYM
      DATA  NCLASS/10/, VSMALL/1.0E-7/, CONST7/0.463/,
     +      LEN/10/, CONST8/1.73205/, CONS11/1.960/, CONS12/2.576/,
     +      IBLANK/'    '/, IASK1/' * '/, IASK2/' **'/
C
C  THE VALUE OF "CONST7" COMES FROM STATISTICAL TABLES FOR
C  TRUNCATED NORMAL DISTRIBUTION IN
C
C      JOHNSON, N.L. AND KOTZ, S.
C      "DISTRIBUTIONS IN STATISTICS: 
C         CONTINUOUS UNIVARIATE DISTRIBUTIONS - 1"
C      PUB. BY WILEY.     NEW YORK, 1970.
C
C  CONST8 = SQRT(3)
C  IVEC(KDIM3)  - FOR STORING ORIGINAL CODE NUMBERS OF OBSERVED DATA,
C         BUT ORDERED ACCORDING TO THE OPTIMUM SEQUENCE ESTABLISHED
C         AND WITH NEGATIVE VALUES WHENEVER AN EVENT IS ON THE SAME
C         LEVEL AS A PRECEDING EVENT.
C  RMAT(KDIM3, 3)
C         COLUMN 1:  CUMULATIVE DISTANCES
C         COLUMN 2:  FIRST ORDER DIFFERENCES
C         COLUMN 3:  SECOND ORDER DIFFERENCES
C  SECDIF(KDIM1 * KDIM3)  IS ALSO FOR SECOND ORDER DIFFERENCES,
C         BUT VALUES ARE STORED AS ONE CONTINUOUS SEQUENCE.
C
      isave=unit
      write(92,10001) ns
10001 format(' NUMBER OF WELLS = ',i4)
10000 DO 10 K = 1,10
          CLASS(K) = 0.0
          CLAS(K) = 0.0
   10 CONTINUE
      DO 20 I = 1,N
          RUNIQ(I,1) = 0.0
          RUNIQ(I,2) = 0.0
   20 CONTINUE
      KK = 0
C
C  MAJOR LOOP TO OBTAIN SECOND ORDER DIFFERENCE STATISTICS
C  AND VALUES FOR "A95" AND "A99"
C
      DO 150 I = 1,NS
          RSTEP = 0.0
          DO 30 J = 1,KDIM3
              ID = IX(I,J)
              IF (ID.GT.0) RSTEP = RSTEP + 1.0
              IF (ID.EQ.0) GO TO 40
   30     CONTINUE
   40     RSTEP = RSTEP - 1.0
          DO 50 K = 1,KDIM3
              RMAT(K,1) = 0.0
              RMAT(K,2) = 0.0
              RMAT(K,3) = 0.0
              IVEC(K) = 0
   50     CONTINUE
          ICNT = 0
          DO 80 J = 1,KDIM3
              ID = IX(I,J)
              IF (ID.EQ.0) GO TO 90
              IAD = IABS (ID)
              ICD = ICODE(IAD)
              DO 60 K = 1,MMAX
                  IF (IRCODE(K).EQ.ICD) GO TO 70
   60         CONTINUE
   70         ID2 = K
              IF (J.EQ.1) IMIN = ID2
              IF (ID.LT.0) ICD = ICD * (-1)
              IVEC(J) = ICD
              RMAT(J,1) = QDAR(ID2)
              ICNT = ICNT + 1
   80     CONTINUE
   90     IMAX = ID2
          ICN = ICNT - 1
C
C      "ICNT" IS THE LENGTH OF THE SEQUENCE
C
          DO 100 J = 1,ICN
              QRAY = RMAT(J+1,1) - RMAT(J,1)
              AVEINC = (QDAR(IMAX) - QDAR(IMIN)) / RSTEP
              IF (IVEC(J+1).GT.0) QRAY = QRAY - AVEINC
              RMAT(J,2) = QRAY
  100     CONTINUE
          DO 110 J = 2,ICN
              JM1 = J - 1
              RMAT(J,3) = RMAT(J,2) - RMAT(JM1,2)
              RMA = RMAT(J,3)
              KK = KK + 1
              SECDIF(KK) = RMA
  110     CONTINUE
  150 CONTINUE
      XXX = FLOAT (KK)
      XXK = XXX * 0.2
      KXX = INT (XXK + 0.5)
      XX = KK - (2 * KXX)
      K = 1
C
  220 BIGPOS = SECDIF(1)
      DO 200 I = 1,KK
          IF (SECDIF(I).GE.BIGPOS)  GO TO 240
              GO TO 200
  240     BIGPOS = SECDIF(I)
              IHIGH = I
  200 CONTINUE
      SECDIF(IHIGH) = VSMALL
C
      BIGNEG = SECDIF(1)
      DO 210 I = 1,KK
          IF (SECDIF(I).LE.BIGNEG)  GO TO 230
              GO TO 210
  230     BIGNEG = SECDIF(I)
              ILOW = I
  210 CONTINUE
      SECDIF(ILOW) = VSMALL
      K = K + 1
      IF (K.LE.KXX)  GO TO 220
C
      S = 0.0
      SS = 0.0
      DO 260 I = 1,KK
          IF (SECDIF(I).EQ.VSMALL)  GO TO 260
          S = S + SECDIF(I)
          SS = SS + (SECDIF(I) ** 2)
  260 CONTINUE
      AVE = S / XX
      VAR = (SS - (XX * AVE * AVE)) / (XX - 1.0)
      SD = SQRT (VAR)
      SDN = SD / CONST7
      RATIO = SDN / CONST8
      CQ = (SDN * SDN) + 1.0
      RHO = 2.0 - SQRT (CQ)
      PR = 1.0/XXX + (2.0 * RHO * (XXX / (1.-RHO) -
     +        1.0 / (1.-RHO)**2) / (XXX * XXX))
      KPR = INT (1.0 / PR + 0.5)
C
      write(unit,2011)
      write(unit,2012)
 2011 format(/////'NORMALITY TEST RESULTS')
 2012 format('______________________')
      write(unit,2013)
 2013 format(/'NOTE: THE NORMALITY TEST IS APPLIED TO SCALING RESULTS'
     +' ONLY')
      WRITE (UNIT,2000)
 2000 FORMAT (////'SECOND ORDER DIFFERENCE STATISTICS'//)
      WRITE (UNIT,2010) KK, AVE, SDN, RATIO, RHO
 2010 FORMAT (1X, 'N = ', I4, '   AVE = ', F7.5, '  SD = ',
     +  F7.5, '   RATIO = ', F4.2, '   RHO = ', F5.3)
      WRITE (UNIT,2020) KPR
 2020 FORMAT (/1X, 'EQUIVALENT NUMBER OF VALUES = ', I4)
      write (unit,2021)
      write (unit,2022)
 2021 format(///'NOTE: PURPOSE OF SECOND ORDER DIFFERENCE STATISTICS',
     +/'IS TO FIT A GAUSSIAN FREQUENCY DISTRIBUTION CURVE TO THE',
     +/'CENTRAL PART OF THE HISTOGRAM OF ALL SECOND ORDER DIFFERENCES')
 2022 FORMAT('IN ALL WELLS. SUCCESSIVE VALUES ARE AUTOCORRELATED (RHO)',
     +/'IMPLYING THAT THE EQUIVALENT NUMBER OF VALUES IS LESS THAN THE',
     +/'ACTUAL NUMBER OF SECOND ORDER DIFFERENCES (N)'//)
      A95 = CONS11 * SDN
      A99 = CONS12 * SDN
C
C
C  MAJOR LOOP TO PERFORM NORMALITY TESTS FOR ALL WELLS
C
      itnt=0
      DO 280 I = 1,NS
          WRITE (UNIT,1010)
          WRITE (UNIT,1020) (NAME(I,J), J = 1,5)
 1010     FORMAT (///'     NORMALITY TEST'//)
 1020     FORMAT (5X, 10A4 ///)
11020     format(1x,' WELL NO.',i3,'   HAS LENGTH = ',i4)
          WRITE (UNIT,1025)
 1025     FORMAT (50X, 'CUM. DIST.  2ND ORDER DIFF.'/)
C
          RSTEP = 0.0
          DO 960 J = 1,KDIM3
              ID = IX(I,J)
              IF (ID.GT.0)  RSTEP = RSTEP + 1.0
              IF (ID.EQ.0)  GO TO 965
  960     CONTINUE
  965     RSTEP = RSTEP - 1.0
          DO 970 K = 1,KDIM3
              RMAT(K,1) = 0.0
              RMAT(K,2) = 0.0
              RMAT(K,3) = 0.0
              IVEC(K) = 0
  970     CONTINUE
          ICNT = 0
          DO 980 J = 1,KDIM3
              ID = IX(I,J)
              IF (ID.EQ.0)  GO TO 990
              IAD = IABS (ID)
              ICD = ICODE(IAD)
              DO 973 K = 1,MMAX
                  IF (IRCODE(K).EQ.ICD)  GO TO 975
  973         CONTINUE
  975         ID2 = K
              IF (J.EQ.1)  IMIN = ID2
              IF (ID.LT.0)  ICD = ICD * (-1)
              IVEC(J) = ICD
              RMAT(J,1) = QDAR(ID2)
              ICNT = ICNT + 1
  980     CONTINUE
  990     IMAX = ID2
          ICN = ICNT - 1
C
C      AGAIN, "ICNT" IS THE LENGTH OF THE SEQUENCE
C
          IF (ICNT.GE.3)  GO TO 992
              WRITE (UNIT,991)
  991         FORMAT (////1X, '(SORRY - RESULTS NOT POSSIBLE)')
              GO TO 280
  992     DO 993 J = 1,ICN
              QRAY = RMAT(J+1, 1) - RMAT(J,1)
              AVEINC = (QDAR(IMAX) - QDAR(IMIN)) / RSTEP
              IF (IVEC(J+1).GT.0)  QRAY = QRAY - AVEINC
              RMAT(J,2) = QRAY
  993     CONTINUE
          DO 996 J = 2,ICN
              JM1 = J - 1
              RMAT(J,3) = RMAT(J,2) - RMAT(JM1,2)
              RMA = RMAT(J,3)
  996     CONTINUE
c          if(unit.ne.92)write(92,11020) i,icn+1
C
          ID = IABS (IVEC(1))
          WRITE (UNIT,1030) (ITITLE(ID,L),L =1,LEN), IVEC(1), RMAT(1,1)
 1030     FORMAT (1X, 10A4, I4, 5X, F10.4)
          DO 120 J = 2,ICN
              RMA = RMAT(J,3)
              AQRAY = ABS (RMA)
              ISYM = IBLANK
              IF (AQRAY.GT.A95) then
              ISYM = IASK1
              itnt=itnt+1
              endif
              IF (AQRAY.GT.A99) ISYM = IASK2
              ID = IABS (IVEC(J))
              WRITE (UNIT,1040) (ITITLE(ID,L), L = 1,LEN), IVEC(J),
     +            RMAT(J,1), RMAT(J,3), ISYM
 1040         FORMAT (1X,10A4,I4,5X,2F10.4,A3)
              ATEMP = RMAT(J,3) / SDN
              CALL ZTOF (ATEMP, BTEMP)
              RMAT(J,3) = BTEMP
  120     CONTINUE
          ID = IABS (IVEC(ICNT))
          WRITE (UNIT,1030) (ITITLE(ID,L), L = 1,LEN), IVEC(ICNT),
     +           RMAT(ICNT,1)
          WRITE (UNIT,1080)
 1080     FORMAT (/10X, '*  -GREATER THAN 95% PROBABILITY THAT',
     +           ' EVENT IS OUT OF POSITION')
          WRITE (UNIT,1090)
 1090     FORMAT (10X, '** -GREATER THAN 99% PROBABILITY THAT',
     +           ' EVENT IS OUT OF POSITION')
          DO 140 K = 1,NCLASS
              DO 130 J = 2,ICN
                  XK = FLOAT (K) / NCLASS
                  IF (RMAT(J,3).LT.XK) CLASS(K) = CLASS(K) + 1.0
  130         CONTINUE
  140     CONTINUE
          IF (INIQ.NE.1) GO TO 280
       IF (NUNIQ(I).EQ.1) CALL XUNIQ2 (N,I,ICNT,IVEC,RMAT,RUNIQ,kuniq)
  280 CONTINUE
      write(unit,1009)
 1009 format(//1x)
      WRITE (UNIT,1010)
      WRITE (UNIT,1070)
 1070 FORMAT (1X, 'COMPARISON OF OBSERVED AND EXPECTED OCCURRENCES OF',
     + ' SECOND ORDER DIFFERENCE VALUES'//)
      WRITE (UNIT,1060)
 1060 FORMAT (1X,' CLASS NO. OBSERVED  EXPECTED DIFFERENCE   DELTA'//)
      TCL = FLOAT (KK) / NCLASS
      CLAS(1) = CLASS(1)
      DO 160 K = 2,NCLASS
          KM1 = K - 1
          CLAS(K) = CLASS(K) - CLASS(KM1)
  160 CONTINUE
      CHI = 0.0
      DO 170 K = 1,NCLASS
          DIFF = CLAS(K) - TCL
          ICL = CLAS(K)
          DELTA = FLOAT (KPR) * DIFF * DIFF / (TCL * XXX)
          CHI = CHI + DELTA
          ISYM = IBLANK
          IF (DELTA.GT.3.84)  ISYM = IASK1
          IF (DELTA.GT.6.63)  ISYM = IASK2
          WRITE (UNIT,1050) K, ICL, TCL, DIFF, DELTA, ISYM
  170 CONTINUE
      ISYM = IBLANK
      IF (CHI.GT.14.1)  ISYM = IASK1
      IF (CHI.GT.18.5)  ISYM = IASK2
      WRITE (UNIT,2030) CHI, ISYM
 2030 FORMAT (//1X, 26X, 'CHI-SQUARED = ', F10.3, A3)
 1050 FORMAT (1X, I6, I10, 2X, 3F10.3, A3)
      WRITE (UNIT,2040)
 2040 FORMAT (/10X, '*  -GREATER THAN 95% PROBABILITY THAT',
     + ' DIFFERENCE IS NOT ZERO')
      WRITE (UNIT,2050)
 2050 FORMAT (10X, '** -GREATER THAN 99% PROBABILITY THAT',
     + ' DIFFERENCE IS NOT ZERO'///)
      IF (INIQ.NE.1.or.unit.eq.92) GO TO 190
      DO 180 I = 1,N
          IF (RUNIQ(I,2).LE.0.0) GO TO 180
          RUNIQ(I,1) = RUNIQ(I,1) / RUNIQ(I,2)
          MMAX = MMAX + 1
          IF (MMAX.GT.KDIM6+MAXUQ)  WRITE (UNIT,2060)
 2060     FORMAT (///1X, '*** ERROR:  SUBSCRIPT OUT OF RANGE IN COMP()',
     +      ' AT ARRAY "QDAR"'/ 16X, 'FURTHER RESULTS MAY BE ERRONEOUS')
          QDAR(MMAX) = RUNIQ(I,1)
          IRCODE(MMAX) = I
  180 CONTINUE
  190 continue
      if(unit.ne.92) then
      unit=92
      goto 10000
      endif
      unit=isave
      RETURN
      END
      SUBROUTINE CYCLE (ICORT,ICYC,WVEC,iqcyc,IA,A,B,CC,IPOS,UNIT,isrt)
C
C ... SUBROUTINE CYCLE IS USED BY RASC TO CORRECT ARTIFICIALLY
C  FOR CYCLICITY IN RANKING SOLUTION.  ALL EVENT PAIRS INVOLVED
C  IN A CYCLE ARE EXAMINED AND THE ELEMENT PAIR WITH THE
C  SMALLEST ABSOLUTE DIFFERENCE IS ZEROED.
C
C  ACCEPTS:  ICORT  - (SWITCH)
C                   = 1,   FOR REPLACEMENT OF ZEROED ELEMENTS
C                    ELSE      SET PAIR OF ELEMENTS TO ZERO
C            WVEC   - (VECTOR OF TYPE "POS")
C            IA     - NUMBER OF CYCLES
C
C  RETURNS:  A,B,CC (300,300,600)
C            IA
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      common /text/ name(kdim1,10), ititle(kdim4,10)
      INTEGER IPOS(kdim6), IQCYC(KDIM5), IRC(KDIM5), t1,t2,t3, UNIT
      integer nss(50),ixdiff(50),min1(10),min2(10)
      REAL    A(300), B(300), CC(600), RDIF(KDIM5), WVEC(KDIM5), t5
      real p(50),sump(50),xn(50),xdiff(50)
C
C  IQCYC(*) - CYCLING EVENTS (ORIGINAL CODE NUMBERS)
C  IRC(*)   - EVENT POSITIONS (BASED ON WVEC(KDIM5))
C  t1, t2, t3, ...   temporary variables (PC-version only) 1985-12-12
      IF (ICORT.EQ.1) GO TO 60
C
C  DETERMINE POSITIONS OF EVENTS INVOLVED IN CYCLE
C
      DO 20 L = 1,MMAX
          RCID = IPOS(L)
          DO 10 J = 1,ICYC
              IF (WVEC(J).EQ.RCID) IRC(J) = L
   10     CONTINUE
   20 CONTINUE
C
C  DETERMINE MINIMUM OF ELEMENT DIFFERENCES
C
      ICYCM1 = ICYC - 1
      DO 30 J = 1,ICYCM1
          t1 = irc(j)
          t2 = irc(j+1)
          xdiff(j)=c(t1,t2)
          if(c(t1,t2).gt.c(t2,t1)) xdiff(j)=c(t2,t1)
          xn(j)=c(t1,t2)+c(t2,t1)
          ixdiff(j)=xdiff(j)
c          RDIF(J) = ABS (C(t1,t2) - C(t2,t1))
   30 CONTINUE
      t1 = irc(icyc)
      t3 = irc(1)
      xdiff(icyc)=c(t1,t3)
      if(c(t1,t3).gt.c(t3,t1)) xdiff(icyc)=c(t3,t1)
      xn(icyc)=c(t1,t3)+c(t3,t1)
      ixdiff(icyc)=xdiff(icyc)
c      do 20207 j=1,icyc
c      write(unit,20208) j,xdiff(j),ixdiff(j),xn(j)
c20207 continue
c20208 format(i2,f5.1,i4,f5.1)
      do 20307 j=1,icyc
c      if(j.eq.2) write(unit,1003) rdif(1)
      nss(j)=xn(j)
      p(1)=0.5**nss(j)
      sump(1)=p(1)
      if(ixdiff(j).eq.0) then
      RDIF(J)=1.0-2.0*sump(1)
      goto 20307
      endif
      do 30307 k=2,nss(j)+1
      p(k)=p(k-1)*(nss(j)-k+2)/(k-1)
      sump(k)=sump(k-1)+p(k)
      if(ixdiff(j).eq.(k-1)) then
      RDIF(j)=1.0-2.0*sump(k)
      goto 20307
      endif
30307 continue
20307 continue
      
c      RDIF(ICYC) = ABS (C(t1,t3) - C(t3,t1))
      DIFMIN = RDIF(1)
      MIN1(1) = 1
      DO 40 J = 2,ICYC
          IF(RDIF(J).GE.DIFMIN) GO TO 40
          DIFMIN = RDIF(J)
          MIN1(1) = J
   40 CONTINUE
      MIN2(1) = MIN1(1) + 1
      IF (MIN1(1).EQ.ICYC) MIN2(1) = 1
      kdex=1
      i=1
      if(min1(1).lt.icyc) then
      ii=min1(1)+1
      do 40307 k=ii,icyc
      sdif=abs(rdif(k)-rdif(min1(1)))
      if(sdif.lt.0.001) then
      i=i+1
      min1(i)=k
      min2(i)=k+1
      if(k.eq.icyc) min2(i)=1
      kdex=kdex+1
      endif
40307 continue
      endif
C
C  PRINT CYCLING EVENTS (ORIGINAL CODE NUMBERS), POSITIONS,
C  AND MATRIX ELEMENTS.
C
      DO 41 J = 1,ICYC
          IQ = IPOS(IRC(J))
          IQCYC(J) = ICODE(IQ)
   41 CONTINUE

C
      WRITE (UNIT,1000) (IQCYC(J), J = 1,ICYC)
c      WRITE (UNIT,1001) (IRC(J), J = 1,ICYC)
      write (unit,71002) (rdif(j), j=1,icyc)
      WRITE (UNIT,1002)
      DO 50 I = 1,ICYC
          WRITE (UNIT,1003) (C(IRC(I),IRC(J)), J = 1,ICYC)
   50 CONTINUE
      do 10307 i=1,kdex
      WRITE (UNIT,1004) iqcyc(MIN1(i)), iqcyc(MIN2(i))
10307 continue

      write (unit,41308)
      do 41307 j=1,icyc
      write(unit,41309) iqcyc(j),(ititle(iqcyc(j),jj), jj=1,10)
41307 continue
41308 format(//10x,  ' NAMES OF EVENTS IN PRECEDING CYCLE:')
41309 format(10x,i6,2x,10a4)


C
C  ZERO THE ELEMENT PAIR WITH THE SMALLEST DIFFERENCE
C
      do 60307 i=1,kdex
      A(IA) = WVEC(MIN1(i))
      B(IA) = WVEC(MIN2(i))
      t1 = irc(min1(i))
      t2 = irc(min2(i))
      CC(IA) = C(t1,t2)
      IA = IA + 1
      CC(IA) = C(t2,t1)
      IA = IA + 1
      C(t1,t2) = 0.0
      C(t2,t1) = 0.0
60307 continue
C
c  new section, July 8, 1995
c
      if(isrt.ne.0.and.isrt.ne.1) then
        do 110 i=1,icyc
          k=i+1
          if(k.gt.icyc) goto 110
          do 111 j=k,icyc
          if(c(irc(i),irc(j)).ge.c(irc(j),irc(i))) goto 111
          itemp=iqcyc(i)
          iqcyc(i)=iqcyc(j)
          iqcyc(j)=itemp
  111     continue
  110 continue
      write(unit,1005) (iqcyc(i),i=1,icyc)
c  112 write(*,*) (iqcyc(i),i=1,icyc)
      endif
 1000 FORMAT (///10X, ' CYCLING EVENTS: ', 9I5)
c 1001 FORMAT (/10X, 'EVENT POSITIONS:', 9I6)
71002 format(/12x,'PROBABILITIES:  ',9f5.2)
 1002 FORMAT (/10X, 'MATRIX ELEMENTS:')
 1003 FORMAT (27X, 9F5.1)
 1004 FORMAT (/'  (',I3,',',I3,')  ZEROED - BASED',
     + ' ON LEAST DIFFERENCE BETWEEN PAIRS')
 1005 FORMAT (/16X,'NEW ORDER:',9I6)
      GO TO 90
C
C  REPLACE ELEMENTS ZEROED IN CORRECTION OF CYCLICITY
C
   60 IF (IA.LE.1)  GO TO 90
      IAM = IA - 2
80307 DO 80 I = 1,IAM,2
          DO 70 J = 1,MMAX
              t5 = float (ipos(j))
              IF (A(I) .EQ. t5)  IA1 = J
              IF (B(I) .EQ. t5)  IB1 = J
   70     CONTINUE
          C(IA1,IB1) = CC(I)
          C(IB1,IA1) = CC(I+1)
c          write (unit, 20209) ia1,ib1,cc(i),cc(i+1)
c20209 format(2i4,2f10.5)
   80 CONTINUE
   90 RETURN
      END
      SUBROUTINE DENDRO (itony,iutem,IPAIR,XLEV,UNIT,st)
C
C ... SUBROUTINE TO PRINT A DENDROGRAM
C
C  ADAPTED FOR RASC FROM "PROGRAM 7.8"    IN
C      DAVIS, J.C.   "STATISTICS AND DATA ANALYSIS IN GEOLOGY"
C      PUB. BY WILEY.    NEW YORK, 1973.
C
C  ACCEPTS:  IPAIR  - (FROM ORDER())
C            XLEV   - (FROM ORDER())
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3),ICODE(KDIM4), IRCODE(KDIM4),
     +C(KDIM6,KDIM6)
      COMMON  /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      COMMON  /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER IPAIR(2,KDIM6+MAXUQ), UNIT,iutem(maxuq,2)
      REAL    XLEV(KDIM6+MAXUQ), XX(13),st(kdim6+maxuq)
      CHARACTER   NAME*4, ITITLE*4, ISTAR*3, ISTAR2*3
      CHARACTER*1  IBLNK, ICI, ICM,IOUT(61) 
      CHARACTER*2  LOGO 
      LOGICAL MARKON, NODON
C
      IBLNK = ' '
      ICI = 'I'
      ICM = '-'
      ISTAR = ' *'
      ISTAR2 = '**'
C
      M2 = MMAX - 1
      WRITE (7,8) MMAX
      if(itony.eq.1.and.iutem(1,1).eq.0) write(71,8)mmax
      if(itony.eq.2.and.iutem(1,1).gt.0) write(71,8)mmax
    8 FORMAT ('***'/, 1X,I10)
C
C  FIND LARGEST AND SMALLEST SIMILARITY CO-EFFICIENT
C
      XMIN = XLEV(1)
      XMAX = XMIN
      kkk=0
      do 9 js=1,m2
      node2=ipair(2,js)
      if(iuniq(node2,1).eq.1.and.js.gt.1.and.kkk.eq.0) then
      xlev(js)=xlev(js)+xlev(js+1)
      kkk=1
      goto 9
      endif
      if(iuniq(node2,1).eq.1.and.js.gt.1.and.kkk.eq.1) then
      xlev(js-1)=xlev(js-1)+xlev(js+1)
      kkk=2
      goto 9
      endif
      if(iuniq(node2,1).eq.1.and.js.gt.1.and.kkk.eq.2)
     + xlev(js-2)=xlev(js-2)+xlev(js+1)
      kkk=0
    9 continue
      DO 10 I = 1,M2
          IF (XLEV(I).LT.XMIN) XMIN = XLEV(I)
          IF (XLEV(I).GT.XMAX) XMAX = XLEV(I)
   10 CONTINUE
      DX = (XMAX - XMIN) / 25.0
      XMIN = XMIN - DX
      XMAX = XMAX + DX
      WRITE (7,11) XMAX,XMIN
      if(itony.eq.1.and.iutem(1,1).eq.0) write(71,11)xmax,xmin
      if(itony.eq.2.and.iutem(1,1).gt.1) write(71,11)xmax,xmin
   11 FORMAT (1X, 2F17.8)
      DX = -(XMAX - XMIN) / 60.0
      XMIN = XMAX
C
C  BLANK OUT PRINT LINE ARRAY
C
      DO 30 I = 1,61
          IOUT(I) = IBLNK
   30 CONTINUE
C
C  PRINT DENDROGRAM
C
      X = XMIN
      DO 40 I = 1,13
          XX(I) = X
          X = X + (DX * 5.0)
   40 CONTINUE
      WRITE (UNIT,2000)
      WRITE (UNIT,2001) (XX(I), I = 2,12,2)
      WRITE (UNIT,2002) (XX(I), I = 1,13,2)
      WRITE (UNIT,2003)
      JS = 1
      jjs=1
      MARKON = .FALSE.
      NODON = .FALSE.
      NODE = IPAIR(1,1)
   50 X = XMIN
c      if(iuniq(node2,1).eq.1.and.js.gt.1)
c     +  xlev(js)=xlev(js)+xlev(js+1)
          IF (JS.NE.MMAX) X = XLEV(JS)
      if(iuniq(node,1).eq.1) goto 61
          IS = IFIX ((X - XMIN) / DX) + 1
          DO 60 I = IS,61
              IOUT(I) = ICM
   60     CONTINUE
   61     LOGO = '  '
          IF (IUNIQ(NODE,2).EQ.1) LOGO = ISTAR
          IF (IUNIQ(NODE,2).EQ.1) MARKON = .TRUE.
          IF (IUNIQ(NODE,1).EQ.1) LOGO = ISTAR2
          IF (IUNIQ(NODE,1).EQ.1) NODON = .TRUE.
          IF (JS.NE.MMAX)
     +    WRITE (UNIT,2004) IOUT,NODE,LOGO,(ITITLE(NODE,J),J=1,10)
          IF (JS.NE.MMAX) then
          WRITE (7,62) NODE,X,LOGO,(ITITLE(NODE,J), J=1,10)
          if(itony.eq.1.and.iutem(1,1).eq.0) then 
          write(71,62) node,x,logo,(ititle(node,j), j=1,10),st(js)
          endif
          if(itony.eq.2.and.iutem(1,1).gt.0) then
          if(logo.ne.istar2) then
          write(71,62) node,x,logo,(ititle(node,j), j=1,10),st(jjs)
          endif
          if(logo.eq.istar2) then
          write(71,62) node,x,logo,(ititle(node,j),j=1,10)
          jjs=jjs-1
          endif
          endif
          endif
   62     FORMAT (1X, I3, F10.4, A3, 10A4,',',f10.4)
          IF (JS.EQ.MMAX)
     +    WRITE (UNIT,2006) IOUT, NODE, LOGO, (ITITLE(NODE,J), J = 1,10)
          IF (JS.EQ.MMAX) then
          WRITE (7,63) NODE,LOGO,(ITITLE(NODE,J), J=1,10)
          if(itony.eq.1.and.iutem(1,1).eq.0) then
          write(71,163) node,logo,(ititle(node,j), j=1,10)
          endif
          if(itony.eq.2.and.iutem(1,1).gt.0) then
          write(71,163) node,logo,(ititle(node,j), j=1,10)
          endif
          itony=itony+1
          endif
   63     FORMAT (1X, I3, A3, 10A4)
  163     format(1x,i3,3x,'-9.9999',a3,10a4,',')
          IF (JS.EQ.MMAX) GO TO 80
          DO 70 I = IS,61
              IOUT(I) = IBLNK
   70     CONTINUE
          IOUT(IS) = ICI
          if(logo.ne.istar2) WRITE (UNIT,2010) (IOUT(I), I = 1,61),x
          if(logo.eq.istar2) write (unit,2011) (iout(i), i = 1,61),x
          NODE = IPAIR(2,JS)
c          node2=ipair(2,js+1)
          JS = JS + 1
          jjs=jjs+1
          GO TO 50
   80 WRITE (UNIT,2003)
      WRITE (UNIT,2002) (XX(I), I = 1,13,2)
      WRITE (UNIT,2001) (XX(I), I = 2,12,2)
      WRITE (UNIT,2005)
      if(markon.and.nodon) then
         write (unit,2009)
         go to 90
      endif
      IF (MARKON) WRITE (UNIT,2007)
      IF (NODON) WRITE (UNIT,2008)
   90 continue
 2007 FORMAT (/3X,'*  INDICATES A MARKER HORIZON')
 2008 FORMAT (/3X,'** INDICATES A UNIQUE (RARE) EVENT')
 2009 format (/3x,'* INDICATES A MARKER HORIZON;  ** INDICATES A',
     + ' UNIQUE (RARE) EVENT')
 2000 FORMAT (//)
 2001 FORMAT (4X,6F10.4)
 2002 FORMAT (1X,f8.4,6F10.4)
 2003 FORMAT (4X,'+',12('----+'))
 2004 FORMAT (4X, 61A1, I3, 8x, 1A3, 10A4)
 2010 format(4x,61a1,3x,f8.4)
 2011 format(4x,61a1,4x,'(',f6.4,')')
 2006 FORMAT (4X, 61A1, I3, 8X, 1A3, 10A4)
 2005 FORMAT (/3X, 'DENDROGRAM -  VALUES ALONG X-AXIS ARE INTER',
     + 'EVENT DISTANCES; VALUES ALONG Y-AXIS ARE DISTANCES',
     + ' BETWEEN'/17x, 'AN EVENT AND ITS SUCCESSOR')
      RETURN
      END
      SUBROUTINE DIST (QDAR,MPAIR,AAA,LLL,UNIT)
C
C ... SUBROUTINE TO COMPUTE INTER-EVENT 'DISTANCES' FOR
C  UNWEIGHTED DIFFERENCES
C
C  THIS IS A SIMPLE TECHNIQUE FOR SCALING RESULTS IN DISTANCES BETWEEN
C  THE FOSSIL EVENTS WHICH ARE CLUSTERED BY MEANS OF SUBROUTINE DENDRO.
C
C  SUBSEQUENTLY, THE DISTANCES ARE RE-CALCULATED BY MEANS OF SUB-
C  ROUTINE WDIST IN WHICH THE FREQUENCIES ARE WEIGHTED ACCORDING TO
C  SAMPLE SIZE.   ALL REMAINING OPTIONS WILL USE THE RESULTS OF WDIST
C  AND NOT THOSE OF THE PRESENT SIMPLE TECHNIQUE (DIST).
C
C  ACCEPTS:  LLL   = 0,  IF RESULTS ARE NOT TO BE PRINTED;
C                     OTHERWISE,  RESULTS WILL BE PRINTED.
C
C  RETURNS:  QDAR  - CUMULATIVE FOSSIL DISTANCES.  SENT TO ORDER()
C            MPAIR - FOR WDIST()
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON N, NS, MMAX, X(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +          C(KDIM6,KDIM6)
      INTEGER   UNIT
      DIMENSION MPAIR(KDIM6+MAXUQ), QDAR(KDIM6+MAXUQ)
      DATA      BOUND/3.0/
C
      QDAR(1) = 0.0
      NMAX = MMAX - 1
      IF (LLL.EQ.0) GO TO 10
          WRITE (UNIT,1000)
          WRITE (UNIT,1010)
          WRITE (UNIT,1020)
 1000     FORMAT (//' UNWEIGHTED DISTANCE ANALYSIS'//)
 1010     FORMAT (2X, 'POSITION   FOSSIL     FOSSIL     CUMULATIVE',
     +     '   SUM DIFF   NO.')
 1020     FORMAT (13X, 'PAIRS     DISTANCE     DISTANCE    Z VALUES',
     +     '   PAIRS'/)
   10 CONTINUE
      DO 100 I = 1,NMAX
          RCONT = 0.0
          SDIFF = 0.0
          ISTP = 0
          K = I - 1
   20     IF (K.EQ.0) GO TO 40
C
C      CASE:  K .LT. I
C             EXAMINE 2 ELEMENTS IN SAME ROW, DIFFERENT COLUMN
C             RCOL1, RCOL2  -  TEMPORARY VARIABLES
C
              RCOL1 = C(K,I)
              RCOL2 = C(K,I+1)
              IF (RCOL1.LT.-BOUND .OR. RCOL1.GT.BOUND) GO TO 30
              IF (RCOL2.LT.-BOUND .OR. RCOL2.GT.BOUND) GO TO 30
              RCONT = RCONT + 1.0
              IF (RCOL1.EQ.AAA .AND. RCOL2.EQ.AAA) ISTP = ISTP + 1
              IF (RCOL1.NE.AAA .OR. RCOL2.NE.AAA) ISTP = 0
              IF (ISTP.EQ.5) RCONT = RCONT - 5.0
              IF (ISTP.EQ.5) GO TO 40
              SDIFF = SDIFF + RCOL2 - RCOL1
   30         K = K - 1
              GO TO 20
C
C  CASE:  K .EQ. I    (SPECIAL CASE)
C
   40     IF (ISTP.GT.0 .AND. ISTP.LT.5) RCONT = RCONT - ISTP
          CAA = C(I,I+1)
          IF (CAA.GE.-AAA .AND. CAA.LE.AAA) SDIFF = SDIFF + CAA
          IF (CAA.GE.-AAA .AND. CAA.LE.AAA) RCONT = RCONT + 1.0
          ISTP = 0
          K = I + 2
C
C  CASE:  K .GT. I
C         DIFFERENT ROW, SAME COLUMN
C
   50     IF (K.GT.MMAX) GO TO 70
               RCOL1 = C(I,K)
               RCOL2 = C(I+1,K)
               IF (RCOL1.LT.-BOUND .OR. RCOL1.GT.BOUND) GO TO 60
               IF (RCOL2.LT.-BOUND .OR. RCOL2.GT.BOUND) GO TO 60
               RCONT = RCONT + 1.0
               IF (RCOL1.EQ.AAA .AND. RCOL2.EQ.AAA) ISTP = ISTP + 1
               IF (RCOL1.NE.AAA .OR. RCOL2.NE.AAA) ISTP = 0
               IF (ISTP.EQ.5) RCONT = RCONT - 5.0
               IF (ISTP.EQ.5) GO TO 70
               RDIFF = RCOL1 - RCOL2
               SDIFF = SDIFF + RDIFF
   60          CONTINUE
               K = K + 1
               GO TO 50
   70     IF (ISTP.GT.0 .AND. ISTP.LT.5) RCONT = RCONT - ISTP
          IF (RCONT.EQ.0.0) GO TO 80
          QDIFF = SDIFF / RCONT
          GO TO 90
   80     QDIFF = 0.0
   90     QDAR(I+1) = QDAR(I) + QDIFF
          MPAIR(I) = INT (RCONT)
          IF (LLL.EQ.0) GO TO 100
          WRITE (UNIT,2000) I, IRCODE(I), IRCODE(I+1), QDIFF, QDAR(I+1),
     +      SDIFF, RCONT
 2000     FORMAT (1X, I8, I5, '-', I3, F11.4, F13.4, F13.2, F7.0)
  100 CONTINUE
      RETURN
      END
      SUBROUTINE ECHO (UNIT)
C
C ... SUBROUTINE TO PRINT OUT ALL THE SEQUENCE DATA
C
C  CALLED IN MAIN PROGRAMME, SUBROUTINES HPFILT AND PRESRT
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS, MMAX, IDATA(KDIM1,KDIM3), MM(KDIM4), M(KDIM4),
     +        RMAT(KDIM6,KDIM6)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER UNIT
      CHARACTER*4  NAME, ITITLE
C
      DO 30 I = 1,NS
          ILIMIT = 0
          DO 10 J = 1,KDIM3
              IF (IDATA(I,J).EQ.0)  GO TO 20
              ILIMIT = ILIMIT + 1
   10     CONTINUE
   20     WRITE (UNIT,1000) (NAME(I,J), J = 1,10)
 1000     FORMAT (///2X, 10A4)
          WRITE (UNIT,2000) (IDATA(I,J), J = 1,ILIMIT)
 2000     FORMAT (20I5)
   30 CONTINUE
      RETURN
      END
      SUBROUTINE HPFILT (IC,IOCR,INIQ,IOMAT,UNIT)
C
C ... SUBROUTINE HPFILT REMOVES FROM THE DATA SET ALL EVENTS WHICH DO
C  DO NOT OCCUR AT LEAST "IOCR" TIMES AND RECODES THE MODIFIED DATA
C  WITH CODE NUMBERS RUNNING FROM 1 TO MMAX
C
C  ACCEPTS:  IC    - ORIGINAL SEQUENCE DATA
C            M     - THE OCCURRENCE TABULATION
C
C  RETURNS:  IC    - NOW, THE FILTERED SEQUENCE DATA
C            ICC   - THE FILTERED, RECODED SEQUENCE DATA
C            MMAX  - NUMBER OF DIFFERENT FOSSILS REMAINING
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS, MMAX, ICC(KDIM1,KDIM3), MM(KDIM4), M(KDIM4),
     +        RMAT(KDIM6,KDIM6)
      COMMON /BETA/ IUNIQ(KDIM4,2), NUNIQ(KDIM1), MUNIQ(KDIM1,MAXUQ*2)
      COMMON /TEXT/ NAME(KDIM1,10), ITITLE(KDIM4,10)
      INTEGER IC(KDIM1,KDIM3), UNIT
      CHARACTER*4  NAME, ITITLE
C
      DO 30 I = 1,NS
          DO 10 J = 1,KDIM3
              ICC(I,J) = 0
   10     CONTINUE
          IF (INIQ.NE.1) GO TO 30
          DO 20 J = 1,40
              MUNIQ(I,J) = 0
   20     CONTINUE
   30 CONTINUE
C
C  ELIMINATE EVENTS WHICH OCCUR FEWER THAN "IOCR" TIMES
C
      DO 60 I = 1,NS
          INDIC = 0
          K = 1
          DO 50 J = 1,KDIM3
              ID = IC(I,J)
              IF (ID.EQ.0) GO TO 60
              IDA = IABS (ID)
              IF (INIQ.EQ.1 .AND. IUNIQ(IDA,1).EQ.1) M(IDA) = 0
              IF (M(IDA).LT.IOCR) GO TO 40
              IF (INDIC.EQ.1 .AND. ID.LT.0) ID = ID * (-1)
              ICC(I,K) = ID
              K = K + 1
              INDIC = 0
              GO TO 50
   40         IF (ID.GT.0 .AND. INDIC.EQ.0) INDIC = 1
   50     CONTINUE
   60 CONTINUE
C     WRITE (UNIT,1000)
C1000 FORMAT (//////'   SEQUENCE DATA MODIFIED TO INCLUDE ONLY')
C     WRITE (UNIT,1010) IOCR
C1010 FORMAT ('   THOSE EVENTS WHICH OCCUR AT LEAST',I3,' TIMES')
C     CALL ECHO (UNIT)
C  SEND ICC TO ECHO() VIA BLANK COMMON
C
C  PREPARE TABLES TO BE USED IN RECODING
C
      DO 70 I = 1,N
          M(I) = 0
          MM(I) = 0
   70 CONTINUE
      DO 90 I = 1,NS
          DO 80 J = 1,KDIM3
              ID = ICC(I,J)
              IF (ID.EQ.0) GO TO 90
              IDA = IABS (ID)
              M(IDA) = M(IDA) + 1
   80     CONTINUE
   90 CONTINUE
      IF (IOMAT.EQ.1) WRITE (UNIT,1020)
 1020 FORMAT (////'   RECODE REFERENCE TABLE, OLD CODE VS. NEW CODE')
C  "M" IS USED AS WORK-AREA FOR THE RECODING TABLE
      K = 1
      DO 100 I = 1,N
          ID = M(I)
          IF (ID.EQ.0) GO TO 100
          M(I) = K
          IF (IOMAT.EQ.1) WRITE (UNIT,1030) I, K
 1030     FORMAT (2I6)
          MM(K) = I
          K = K + 1
  100 CONTINUE
      IF (INIQ.NE.1) GO TO 120
      DO 110 I = 1,NS
          IF (NUNIQ(I).EQ.1) CALL XUNIQ1 (I,IC)
  110 CONTINUE
C
C     MAKE A COPY OF THE FILTERED DATA SET BEFORE RECODING.
C          (DESTROY THE UN-FILTERED DATA SET)
C
  120 DO 128 I = 1,NS
          DO 125 J = 1,KDIM3
              IC(I,J) = ICC(I,J)
  125     CONTINUE
  128 CONTINUE
C
C  PERFORM RECODING
C
      MMAX = K - 1
      DO 140 I = 1,NS
          DO 130 J = 1,KDIM3
              ID = ICC(I,J)
              IF (ID.EQ.0) GO TO 140
              IDA = IABS (ID)
              IDD = M(IDA)
              IF (ID.LT.0) IDD = IDD * (-1)
              ICC(I,J) = IDD
  130     CONTINUE
  140 CONTINUE
C     WRITE (UNIT,1040)
C1040 FORMAT (////'   RECODED SEQUENCE DATA')
C  AGAIN, SEND ICC TO ECHO()
C     CALL ECHO (UNIT)
      IF (IOMAT.NE.1)  GO TO 175
          WRITE (UNIT,1050)
 1050     FORMAT (///'   CROSS REFERENCE TABLE, NEW CODE VS. OLD CODE')
C  "MM" IS A SERIAL LIST OF OLD CODE NUMBERS
          DO 160 I = 1,MMAX
              WRITE (UNIT,1060) I, MM(I)
 1060         FORMAT (2I6)
  160     CONTINUE
C 170 WRITE (UNIT,1070) MMAX
C1070 FORMAT (////'    NUMBER OF FOSSILS RETAINED:',I5)
  175 WRITE (7,1071) MMAX
 1071 FORMAT (I10)
C
      IF (MMAX.GT.1 .AND. MMAX.LT.KDIM6)  GO TO 999
      IF (MMAX.GT.KDIM6)  GO TO 220
          WRITE (UNIT,1075)
 1075     FORMAT (///1X, '*** ERROR:  TOO MUCH FILTERING DONE --',
     +      ' NO DATA LEFT'/16X, 'RECOMMENDED ACTION: DECREASE',
     +      ' VALUE OF "IOCR"')
          WRITE (UNIT,1090)
          STOP
  220 WRITE (UNIT,1080) KDIM6
 1080     FORMAT (///1X,'*** ERROR:   TOO MANY EVENTS REMAINING  (>',I4,
     +      ')', /16X, 'CAUSE:  NOT ENOUGH FILTERING DONE',
     +      /16X, 'RECOMMENDED ACTION:  INCREASE VALUE OF "IOCR" OR',
     +      /36X, 'USE A BIGGER VERSION OF RASC')
          WRITE (UNIT,1090)
 1090     FORMAT (///' *** EXECUTION TERMINATED IN   SUBROUTINE HPFILT')
          STOP
C
  999 RETURN
      END
      SUBROUTINE NORMZ (AAA,LLL,IOMAT,UNIT)
C
C ... SUBROUTINE TO COMPUTE 'Z' (NORMAL) VALUES OF FREQUENCIES
C
C      C(I,J)
C  ---------------     FROM CUMULATIVE ORDER MATRIX.
C  C(I,J) + C(J,I)
C
C  ACCEPTS:  C  - CUMULATIVE ORDER MATRIX
C  RETURNS:  UPPER TRIANGLE OF "C", EXCLUDING THE MAIN DIAGONAL
C
C  CALLED IN MAIN PROGRAMME.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +       C(KDIM6,KDIM6)
      INTEGER UNIT
      DATA  CONST/9.0/
      call ztof(-aaa,taill)
      call ztof(aaa,tailr)
C
      DO 40 I = 1,MMAX
          K = I + 1
          DO 30 J = K,MMAX
              RCL1 = C(I,J)
              RCL2 = C(J,I)
              SRCL = RCL1 + RCL2
              IF (SRCL.EQ.0.0) GO TO 20
              QRCL = RCL1 / SRCL
              IF (QRCL.LE.0.0 .OR. QRCL.GE.1.0)  GO TO 10
              CALL FTOZ (QRCL,QX)
              C(I,J) = QX
              GO TO 30
   10         IF (QRCL.EQ.1.0) C(I,J) = AAA
              IF (QRCL.EQ.1.0) C(J,I) = TAILL * SRCL
              IF (QRCL.EQ.0.0) C(I,J) = -AAA
              IF (QRCL.EQ.0.0) C(J,I) = TAILR * SRCL
              GO TO 30
   20         C(I,J) = CONST
   30     CONTINUE
   40 CONTINUE
      IF (LLL.EQ.0 .OR. IOMAT.NE.1)  GO TO 60
      WRITE (UNIT,1000)
 1000 FORMAT (////'  UPPER TRIANGLE OF NORMAL Z VALUES')
      DO 50 I = 1,MMAX
          WRITE (UNIT,1001)
 1001     FORMAT (///)
          WRITE (UNIT,1002) (C(I,J), J = 1,MMAX)
 1002     FORMAT (1X, 15F8.3)
   50 CONTINUE
   60 RETURN
      END
      SUBROUTINE OCCTAB (UNIT,iocr)
C
C ... SUBROUTINE TO PRINT A SUMMARY TABLE SHOWING HOW MANY TIMES AN
C  EVENT APPEARS IN THE DATA SET AS A WHOLE.
C  ACCEPTS:  M(KDIM4)  -  TALLY OF DICTIONARY EVENTS  (FROM M/PROG)
C
C  CALLED IN MAIN PROGRAMME.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS, MMAX, ICC(KDIM1,KDIM3), ICODE(KDIM4), M(KDIM4),
     +        RMAT(KDIM6,KDIM6)
      common/text/name(kdim1,10),ititle(kdim4,10)
      INTEGER ICOL(10), MCUM(KDIM1), MM(KDIM1), UNIT,MN,
     +        IOC1(KDIM4), IOC2(KDIM4)
      character*4 ititle
      DATA    NCOLS/6/
C
      WRITE (UNIT,1000)
      WRITE (UNIT,2000)
      WRITE (UNIT,3000)
 1000 FORMAT (////'TABULATION OF EVENT RECORD OCCURRENCES')
 2000 FORMAT ('DICTIONARY CODE NUMBER VERSUS FREQUENCY OF OCCURRENCE')
 3000 FORMAT (//)
      MN = 0
      DO 2 I=1,N
      IF (M(I).NE.0) THEN
      MN = MN + 1
      IOC1(MN) = I
      IOC2(MN) = M(I)
      ENDIF
    2 CONTINUE
C
C     NROWS = SMALLEST INTEGER .GE. (MN / FLOAT(NCOLS))
      NROWS = INT ((MN / FLOAT (NCOLS)) + 0.99)
      II = 1
    3 ICOL(NCOLS) = (NCOLS - 1) * NROWS + II
      IF (ICOL(NCOLS).GT.MN)  GO TO 10
          DO 5 IND = 1,NCOLS
              ICOL(IND) = (IND - 1) * NROWS + II
    5     CONTINUE
          WRITE (UNIT,4000) (IOC1(ICOL(K)), IOC2(ICOL(K)), K = 1,NCOLS)
 4000     FORMAT (9('  *  ', 2I5))
          II = II + 1
          GO TO 3
   10 CONTINUE
      IF (II.GT.NROWS)  GO TO 17
          DO 15 J = II,NROWS
              NEWCOL = NCOLS - 1
              DO 13 IND = 1,NEWCOL
                  ICOL(IND) = (IND - 1) * NROWS + J
   13         CONTINUE
              WRITE (UNIT,4000) (IOC1(ICOL(K)), 
     +               IOC2(ICOL(K)), K = 1,NEWCOL)
   15     CONTINUE
C
   17 DO 20 I = 1,NS
          MM(I) = 0
   20 CONTINUE
      DO 35 J = 1,N
         IF (M(J).GT.NS)  GO TO 8000
         IF (M(J).GT.0)  MM(M(J)) = MM(M(J)) + 1
   35 CONTINUE
      MCUM(NS) = MM(NS)
C
C  CUMULATE EVENT SEQUENCE
C
      DO 50 J = 2,NS
          I = NS + 1 - J
          MCUM(I) = MCUM(I+1) + MM(I)
   50 CONTINUE
      WRITE (UNIT,1000)
C LET IEND = SMALLEST INTEGER .GE. (NS/24.0)
      IEND = INT (NS / 24.0 + 0.99)
      nuni=0
      nuni2=0
      do 54 i=1,n
      if(m(i).lt.iocr.and.m(i).gt.0) numi=numi+1
      if(m(i).gt.0) numi2=numi2+1
   54 continue
      write(88,1006) numi
      write(94,1006) numi2
      do 55 i=1,n
      if(m(i).gt.0.)write(94,1005) i,m(i),(ititle(i,j),j=1,10)
      if(m(i).lt.iocr.and.m(i).gt.0)then
      write(88,1005) i,m(i),(ititle(i,j),j=1,10)
      endif
   55 continue
 1005 format(2i4,2x,10a4)
 1006 format(' NUMBER OF EVENTS = ',i4)
      DO 60 I = 1,IEND
          JSTART = (24 * I) - 23
          JEND = NS
          IF (I.LT.IEND)  JEND = JSTART + 23
          WRITE (UNIT,5000) (J, J = JSTART,JEND)
          write(70,5000)(j,j=jstart,jend)
 5000     FORMAT (//3X, 'NUMBER OF WELLS  ', 24I4)
C          WRITE (UNIT,6000) (MM(J), J = JSTART,JEND)
C 6000     FORMAT (3X, 'NUMBER OF EVENTS ', 24I4)
          WRITE (UNIT,7000) (MCUM(J), J = JSTART,JEND)
          write(70,7000) (mcum(j),j=jstart,jend)
 7000     FORMAT (3X, 'CUMULATIVE NUMBER', 24I4)
   60 CONTINUE
      RETURN
C
 8000 WRITE (UNIT,8200) J, M(J)
 8200 FORMAT (//' *** ERROR -  TOO MANY OCCURRENCES.',
     +  /15X, 'FOSSIL',I4,' OCCURS',I4,' TIMES, WHICH IS MORE THAN',
     +  ' THE NUMBER OF WELLS IN THE DATA SET')
      WRITE (UNIT,8210)
 8210 FORMAT (/' NOTE THAT A FOSSIL MUST NOT BE RECORDED MORE THAN',
     +  ' ONCE IN ANY WELL')
      WRITE (UNIT,8220)
 8220 FORMAT (///' *** EXECUTION TERMINATED IN   SUBROUTINE OCCTAB')
      STOP
C
      END
      SUBROUTINE ORDER (QDAR,IPAIR,XLEV,LLL,LOUT,UNIT)
C
C ... SUBROUTINE ORDER IS A SUPPORT ROUTINE FOR DISTANCE CALCULATIONS
C  TO RE-ORDER THE OPTIMUM SEQUENCE ON THE BASIS OF CUMULATIVE
C  INTER-EVENT "DISTANCES".
C
C  ACCEPTS:  MMAX   - NUMBER OF EVENTS TO BE CONSIDERED
C                     (NUMBER OF FOSSILS REMAINING AFTER FILTERING)
C            QDAR   - CUMULATIVE INTERFOSSIL DISTANCES.
C                     FROM DIST(), WDIST(), AND COMP()
C            IRCODE - THE CURRENT OPTIMUM SEQUENCE
C
C  RETURNS:  QDAR   - SORTED IN ASCENDING ORDER
C            IRCODE - OPTIMUM SEQUENCE RE-ORDERED ON THE BASIS OF
C                       ASCENDING QDAR (THE CUMULATIVE DISTANCES)
C            IPAIR(2,*)  - (FOR USE IN DENDRO())
C            XLEV(*)     - (FOR USE IN DENDRO())
C
C  CALLED IN MAIN PROGRAMME.   (MANY TIMES NEAR THE END)
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      PARAMETER (MAXCYC=300, MAXUQ=50)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      INTEGER IPAIR(2,KDIM6+MAXUQ),UNIT
      REAL    QDAR(KDIM6+MAXUQ), XLEV(KDIM6+MAXUQ)
C
C  SORT CUMULATIVE 'DISTANCES' IN ASCENDING ORDER
C
      NMAX = MMAX - 1
      QDAR(1) = 0.0
   10 KFLAG = 0
      DO 20 I = 1,NMAX
          SORT1 = QDAR(I)
          SORT2 = QDAR(I+1)
          IF (SORT1.LE.SORT2) GO TO 20
              QDAR(I) = SORT2
              QDAR(I+1) = SORT1
              ITEMP = IRCODE(I)
              IRCODE(I) = IRCODE(I+1)
              IRCODE(I+1) = ITEMP
              KFLAG = 1
   20 CONTINUE
      IF (KFLAG.GE.1) GO TO 10
      DO 30 I = 1,NMAX
          IPAIR(1,I) = IRCODE(I)
          IPAIR(2,I) = IRCODE(I+1)
          XLEV(I) = QDAR(I+1) - QDAR(I)
   30 CONTINUE
      XLEV(MMAX) = 0.0
      IF (LLL.EQ.0) GO TO 50
          WRITE (UNIT,1000)
          IF (LOUT.EQ.1) WRITE (UNIT,1010)
 1010         FORMAT (1X, 'NOTE:  IN ORDER TO RECALCULATE STANDARD',
     +        ' DEVIATIONS AFTER SORTING,',/' DISTANCE VALUES MUST BE',
     +        ' RECALCULATED'//)
          WRITE (UNIT,1001)
          WRITE (UNIT,1002)
          WRITE (UNIT,1003)
 1000     FORMAT(///1X,'EVENTS ARE SORTED ON THE BASIS OF CUMULATIVE ',
     +    'DISTANCE',/' TO OBTAIN ONLY POSITIVE INTEREVENT DISTANCES'/)
 1001     FORMAT (' NEW        DISTANCE      EVENT     INTER')
 1002     FORMAT (' SEQUENCE   FROM 1ST      PAIRS     EVENT')
 1003     FORMAT ('            POSITION                DISTANCE'/)
          DO 40 I = 1,NMAX
              WRITE (UNIT,1004) IRCODE(I),QDAR(I),IRCODE(I),IRCODE(I+1),
     +          QDAR(I+1)-QDAR(I)
 1004         FORMAT (1X, I4, F13.4, I10, '-', I3, F10.4)
   40     CONTINUE
          WRITE (UNIT,1005) IRCODE(MMAX), QDAR(MMAX)
 1005     FORMAT (1X, I4, F13.4)
   50     RETURN
      END
      SUBROUTINE PRESRT (IOMAT,UNIT)
C
C ... SUBROUTINE FOR PRELIMINARY SEQUENCING OF THE DATA IN ORDER TO
C  OBTAIN AN OPTIMIZED STARTING SEQUENCE
C
C  ACCEPTS:  MMAX  - NUMBER OF DIFFERENT FOSSILS REMAINING AFTER
C                      FILTERING.  (FROM HPFILT())
C            ICC   - FILTERED, RECODED SEQUENCE DATA
C  RETURNS:  RMAT
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON N, NS, MMAX, ICC(KDIM1,KDIM3), MM(KDIM4), IGAD(KDIM4),
     +       RMAT(KDIM6,KDIM6)
      DIMENSION MP(KDIM6,3), SCORE(KDIM6,2)
      INTEGER UNIT
C
C  CREATE AN ORDER RELATION MATRIX
C
C     WRITE (UNIT,1000)
C1000 FORMAT (////' PRESORT OPTION INITIATED')
      IF (IOMAT.EQ.1) WRITE (UNIT,1010)
 1010 FORMAT (///' CUMULATIVE ORDER MATRIX'///)
      DO 20 I = 1,MMAX
          DO 10 J = 1,MMAX
              RMAT(I,J) = 0.0
   10     CONTINUE
   20 CONTINUE
C
C  ITEST   = 0,  IF IKK IS FOUND ON THE SAME LEVEL OF THE WELL AS
C                 IAA  (OUR REFERENCE FOSSIL)
C
      DO 60 L = 1,NS
          DO 50 J = 1,MMAX
              ITEST = 0
              K = J
              ID = ICC(L,K)
              IF (ID.EQ.0) GO TO 60
              I = IABS (ID)
   30         K = K + 1
              IF (K.GT.KDIM3)  GO TO 50
              IAA = ICC(L,K)
              IF (IAA.EQ.0) GO TO 50
              IKK = IABS (IAA)
              IF (IAA.LT.0 .AND. ITEST.LE.0) GO TO 40
                  ITEST = 1
                  RMAT(I,IKK) = RMAT(I,IKK) + 1.0
                  GO TO 30
   40         RMAT(I,IKK) = RMAT(I,IKK) + 0.5
                  RMAT(IKK,I) = RMAT(IKK,I) + 0.5
                  GO TO 30
   50     CONTINUE
   60 CONTINUE
      DO 70 I = 1,MMAX
          SCORE(I,2) = 0.0
          IF (IOMAT.EQ.1) WRITE (UNIT,1020) (RMAT(I,J), J = 1,MMAX)
 1020     FORMAT (20F5.1)
   70 CONTINUE
C
C  CALCULATE FOSSIL SCORES
C
      DO 90 I = 1,MMAX
          RCOUNT = 0.0
          DO 80 J = 1,MMAX
              IF (I.EQ.J) GO TO 80
              REL1 = RMAT(I,J)
              REL2 = RMAT(J,I)
              IF (REL1.EQ.0.0 .AND. REL2.EQ.0.0) GO TO 80
              RCOUNT = RCOUNT + 1.0
              FREQ = REL1 / (REL1 + REL2)
              IF (FREQ.LT.0.5) FREQ = 0.0
              IF (FREQ.GT.0.5) FREQ = 1.0
              SCORE(I,2) = SCORE(I,2) + FREQ
   80     CONTINUE
          SCORE(I,2) = SCORE(I,2) * (MMAX - 1) / RCOUNT
   90 CONTINUE
      IF (IOMAT.EQ.1) WRITE (UNIT,1030)
 1030 FORMAT (///' CODE FOLLOWED BY SCORE')
C
C  SCORE(*,2)  -  FIRST COLUMN: (ESSENTIALLY INTEGER) THE POSITIONS
C                SECOND COLUMN: THE PRESORTING SCORES
C
      Z = 1.0
      DO 100 I = 1,MMAX
          SCORE(I,1) = Z
          Z = Z + 1
          IF (IOMAT.EQ.1) WRITE (UNIT,1040) SCORE(I,1), SCORE(I,2)
 1040     FORMAT (2F10.1)
  100 CONTINUE
C
C  SORT THE SCORES IN DESCENDING ORDER
C
      MM1 = MMAX - 1
      DO 120 I = 1,MM1
          MM2 = I + 1
          DO 110 J = MM2,MMAX
              IF (SCORE(I,2).GE.SCORE(J,2)) GO TO 110
              TEMP2 = SCORE(I,2)
              SCORE(I,2) = SCORE(J,2)
              SCORE(J,2) = TEMP2
              TEMP1 = SCORE(I,1)
              SCORE(I,1) = SCORE(J,1)
              SCORE(J,1) = TEMP1
  110     CONTINUE
  120 CONTINUE
      IF (IOMAT.EQ.1) WRITE (UNIT,1050)
 1050 FORMAT (///' SCORES IN DESCENDING ORDER')
      DO 130 I = 1,MMAX
          IF (IOMAT.EQ.1) WRITE (UNIT,1040) SCORE(I,1), SCORE(I,2)
  130 CONTINUE
C
C
      IF (IOMAT.EQ.1) WRITE (UNIT,1060)
 1060 FORMAT (///' NEW CROSS REFERENCE TABLE')
      DO 140 I = 1,MMAX
          MP(I,1) = I
          MP(I,2) = MM(I)
          MP(I,3) = INT (SCORE(I,1))
          IF (IOMAT.EQ.1) WRITE (UNIT,1070) MP(I,1), MP(I,3)
 1070     FORMAT (1X, 2I6)
  140 CONTINUE
C
C  CONSTRUCT A NEW CODE REFERENCE TABLE
C
      DO 150 I = 1,MMAX
          ID = MP(I,3)
          MM(I) = MP(ID,2)
  150 CONTINUE
C
C  SORT THE NEW ORDER
C
      DO 180 I = 1,MM1
          MM2 = I + 1
          DO 170 J = MM2,MMAX
              IF (MP(J,3).GT.MP(I,3)) GO TO 170
              DO 160 K = 1,3
                  ITEMP = MP(I,K)
                  MP(I,K) = MP(J,K)
                  MP(J,K) = ITEMP
  160         CONTINUE
  170     CONTINUE
  180 CONTINUE
C
C  RECODE DATA IN PRELIMINARY SEQUENCE
C
      DO 200 I = 1,NS
          DO 190 J = 1,KDIM3
              ID = ICC(I,J)
              IF (ID.EQ.0) GO TO 200
              IDD = MP(IABS(ID), 1)
              IF (ID.LT.0) IDD = IDD * (-1)
              ICC(I,J) = IDD
  190     CONTINUE
  200 CONTINUE
C     WRITE (UNIT,3000)
C3000 FORMAT (////'   RECODED AND PRESORTED DATA SET')
C     CALL ECHO (UNIT)
      RETURN
      END
      SUBROUTINE REORD (AAA,IPOS)
C
C ... SUBROUTINE TO REORDER THE CUMULATIVE ORDER MATRIX DURING THE
C  FINAL REORDERING OPTION
C
C  CALLED IN MAIN PROGRAMME.
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      INTEGER IPOS(KDIM6)
      DATA  BOUND/3.0/
      call ztof(-aaa,taill)
      call ztof(aaa,tailr)
C
C  RECOMPUTE FREQUENCIES FROM "Z" VALUES
C
      DO 50 I = 1,MMAX
          K = I + 1
          DO 40 J = K,MMAX
              ZED = C(I,J)
              IF (ZED.GE.BOUND) GO TO 10
                  IF (ZED.EQ.AAA) GO TO 20
                  IF (ZED.EQ.-AAA) GO TO 30
                  CALL ZTOF (ZED,FREQ)
                  C(I,J) = C(J,I) * FREQ / (1.0 - FREQ)
                  GO TO 40
   10         CONTINUE
                  C(I,J) = 0.0
                  C(J,I) = 0.0
                  GO TO 40
   20         CONTINUE
                  C(I,J) = C(J,I) / TAILL
                  C(J,I) = 0.0
                  GO TO 40
   30         CONTINUE
                  C(J,I) = C(J,I) / TAILR
                  C(I,J) = 0.0
   40     CONTINUE
   50 CONTINUE
C
C  REORDER THE FINAL RELATION MATRIX
C
      DO 120 I = 1,MMAX
          IFOS = IRCODE(I)
          DO 60 L = 1,MMAX
              IF (ICODE(L).EQ.IFOS) GO TO 70
   60     CONTINUE
   70     INEW = L
          DO 80 K = 1,MMAX
              IAB = IPOS(K)
              IF (IAB.EQ.INEW) GO TO 90
   80     CONTINUE
   90     II2 = K
          IF (I.EQ.II2) GO TO 120
          DO 100 L = 1,MMAX
              TEMP = C(I,L)
              C(I,L) = C(II2,L)
              C(II2,L) = TEMP
  100     CONTINUE
          ITEMP = IPOS(I)
          IPOS(I) = IPOS(II2)
          IPOS(II2) = ITEMP
          DO 110 L = 1,MMAX
              TEMP = C(L,I)
              C(L,I) = C(L,II2)
              C(L,II2) = TEMP
  110     CONTINUE
  120 CONTINUE
      RETURN
      END
      SUBROUTINE SCATTR (IDATA, NST, IWELL,ircodeo,ircodea,UNIT,ifit,
     +itsa)
C
C ... SUBROUTINE TO PRODUCE A SCATTERGRAM FOR A GIVEN WELL
C
C  ACCEPTS:  IDATA  -  THE SEQUENCE DATA FOR THE SCATTERGRAM
C            NST    -  NUMBER OF NON-ZERO ENTRIES IN THE SEQUENCE
C            IWELL  -  THE WELL NUMBER
C
C  ADDED BY     M. HELLER CONSULTANTS
C               HALIFAX, N.S.
C
C  CALLED IN MAIN PROGRAMME
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3), ICODE(KDIM4), IRCODE(KDIM4),
     +        C(KDIM6,KDIM6)
      INTEGER I, IEVENT, IFLAG, J, LEVNUM, MAXCNT, NLEVS, NSAME,
     +        NSTP1, IDATA(KDIM3), ITABLE(50,20), NCOUNT(50), UNIT,
     +        ircodeo(kdim4),ircodea(kdim4)
C
C  MAXLEV  - MAXIMUM NUMBER OF LEVELS (NOT FOSSILS) IN THE WELL
C  MAXSAM  - MAXIMUM NUMBER OF FOSSILS AT THE  *SAME*  LEVEL
C  LEVNUM  - THE LEVEL NUMBER  (THAT WE ARE CURRENTLY EXAMINING)
C
      itemp=unit
      unit=64
      if(ifit.eq.1) unit=65
      MAXLEV = 30
      MAXSAM = 20
      LEVNUM = 0
      NSTP1 = NST + 1
      IFLAG = 0
      J = 1
C
C  MAIN LOOP:  EXAMINE THE ENTIRE SEQUENCE, ONE EVENT AT A TIME
C
      IEVENT = IDATA(J)
 2500 IF (IEVENT.EQ.0 .OR. J.GT.NSTP1 .OR. IFLAG.NE.0) GO TO 4000
          J = J + 1
          IF (IEVENT.LT.0)  GO TO 3100
C
C             CASE:  POSITIVE EVENT-NUMBER (IMPLIES NEW LEVEL)
C                   THIS IS ALWAYS THE CASE THE FIRST TIME THROUGH
C
              IF (LEVNUM.GT.0)  NCOUNT(LEVNUM) = NSAME
              LEVNUM = LEVNUM + 1
              IF (LEVNUM.GT.MAXLEV)  GO TO 3000
                  NSAME = 1
                  ITABLE(LEVNUM,1) = IEVENT
                  GO TO 3200
C
 3000         IFLAG = 1
                  GO TO 3200
C
C             CASE:  NEGATIVE EVENT NR. (THUS AT THE SAME LEVEL)
C
 3100         NSAME = NSAME + 1
              IF (NSAME.GT.MAXSAM)  GO TO 3150
                  ITABLE(LEVNUM,NSAME) = -IEVENT
                  GO TO 3200
C
 3150         IFLAG = 2
                  GO TO 3200
C
 3200         IEVENT = IDATA(J)
              GO TO 2500
C             END-OF-MAIN-LOOP
C
 4000 NLEVS = LEVNUM
      IF (IEVENT.EQ.0 .AND. IFLAG.EQ.0)  GO TO 6000
C
C             CASE:  EXIT CAUSED BY EXCEEDING A MAXIMUM
C
              IF (IFLAG.EQ.2)  GO TO 4200
                  NLEVS = MAXLEV
                  WRITE (UNIT,4100) IWELL, MAXLEV
 4100             FORMAT (1X, '** SCATTERGRAM LIMITS EXCEEDED',
     +             6X, 'AT WELL NUMBER', I4 /6X, 'MORE THAN',
     +             I4, ' DISTINCT LEVELS'/)
                  GO TO 4500
 4200         NSAME = MAXSAM
                  WRITE (UNIT,4300) IWELL, MAXSAM
 4300             FORMAT (1X, '** SCATTERGRAM LIMITS EXCEEDED',
     +             6X, 'AT WELL NUMBER', I4 /6X, 'MORE THAN',
     +             I4, ' COEVAL EVENTS'/)
 4500         NCOUNT(NLEVS) = NSAME
c              DO 4800 I = 1,NLEVS
c                  JLIMIT = NCOUNT(I)
c                  WRITE (UNIT,4700) (ITABLE(I,J), J = 1,JLIMIT)
c 4700             FORMAT (1X, 10I4)
c 4800         CONTINUE
      write(unit,4900)
 4900 format(/1x)
C

C     ELSE  EXIT OCCURRED AT END OF SEQUENCE
C
 6000 NCOUNT(NLEVS) = NSAME
C
C     NOW FIND THE MAXIMUM ON NCOUNT(I)
C
      MAXCNT = 0
      DO 6100 I = 1,NLEVS
          IF (NCOUNT(I).GT.MAXCNT)  MAXCNT = NCOUNT(I)
 6100 CONTINUE
      unit=itemp
C
      CALL SCDRAW(ITABLE,NLEVS,MAXCNT,NCOUNT,ircodeo,ircodea,UNIT,ifit,
     +itsa)
C
      RETURN
      END
      SUBROUTINE SCDRAW (ITABLE,NLEVS,MAXCNT,NCOUNT,ircodeo,ircodea,
     +UNIT,ifit,itsa)
C
C ... SUBROUTINE TO PRINT SCATTERGRAM FOR SUBROUTINE SCATTR
C
C  ADDED BY     M. HELLER CONSULTANTS
C               HALIFAX, N.S.
C               modified by F. P. Agterberg
c
C  THE MEMORY-TO-MEMORY "WRITE" STATEMENT IN LOOP 200 CONVERTS INTEGERS
C  ("I" FORMAT) TO ALPHANUMERICS ("A" FORMAT)
C
C  CALLED BY SUBROUTINE SCATTR
C
      PARAMETER (KDIM1=100, KDIM3=300, KDIM4=998,KDIM5=10, KDIM6=200)
      COMMON  N, NS,MMAX, IX(KDIM1,KDIM3),ICODE(KDIM4), IRCODE(KDIM4),
     +C(KDIM6,KDIM6)
      INTEGER ITABLE(50,20), NCOUNT(50), UNIT,ircodeo(kdim4),
     + ircodea(kdim4)
      CHARACTER*4  BLANK, DASH, MARK, COLOUT(30),marko,marka
      DATA  MARK/' OO '/, BLANK/'    '/, DASH/'----'/,marko/' XX '/,
     + marka/'AAAA'/
      itemp=unit
      unit=64
      if(ifit.eq.1) unit=65
C
      DO 100 I = 1,MAXCNT
          LIMUP = MAXCNT + 1 - I
          DO 200 J = 1,NLEVS
              COLOUT(J) = BLANK
              IF (NCOUNT(J).LT.LIMUP)  GO TO 200
              WRITE (COLOUT(J), '(I4)') ITABLE(J,LIMUP)
  200     CONTINUE
          WRITE (UNIT,400) (COLOUT(J), J = 1,NLEVS)
  400     FORMAT (8X, 1HI, 50A4)
  100 CONTINUE
      WRITE (UNIT,500) (DASH,I = 1,NLEVS)
  500 FORMAT (9H -------I,50A4)
c      itsa=0
      DO 600 I = 1,MMAX
          DO 700 J = 1,NLEVS
              COLOUT(J) = BLANK
              KLIM = NCOUNT(J)
              DO 800 K = 1,KLIM
                  IF (IRCODE(I).EQ.ITABLE(J,K)) COLOUT(J) = MARK
                  if(ircodeo(i).eq.itable(j,k)) colout(j) = marko
                  if(ircodea(i).eq.itable(j,k)) then
                  colout(j) = marka
                  itsa=itsa+1
                  endif
  800         CONTINUE
  700     CONTINUE
          WRITE (UNIT,900) IRCODE(I), (COLOUT(J), J = 1,NLEVS)
  900     FORMAT (2X, I5, ' I ', 50A4)
  600 CONTINUE
      write(unit,1000)
 1000 format(//10x,' OO - EVENTS CLOSEST TO LINE OF CORRELATION',/
     +10X,' XX - EVENTS WHICH ARE RELATIVELY CLOSE',/
     +9X,'AAAA - EVENTS WHICH MAY BE OUT OF PLACE (95% PROBABILITY)'/)
      unit=itemp
      RETURN
      END
      subroutine fit(idata,af1,bf1,cf1,nst,qdar,xdev,sdev,icasc,ifit)
c
c ... subroutine to fit lines of correlation to scattergram data
c
c  accepts:  idata  -  sequence data
c            nst    -  number of non-zero entries in the sequence
c            qdar   -  RASC distances
c
c  returns:   af1, etc. -  estimated values for this well
c             xdev    -  deviations for this well
c             sdev   -  standard deviation
c
c  added by F.P. Agterberg, 10 August, 1995
c
c  called in main programme
c
      parameter (KDIM1=100,KDIM3=300,kdim4=998,kdim5=10,KDIM6=200)
      parameter (MAXUQ=50)
      common  n,ns,mmax,ix(kdim1,kdim3),icode(kdim4),ircode(kdim4),
     +        c(kdim6,kdim6)
      common /text/ name(kdim1,10),ititle(kdim4,10)
      integer idata(kdim3)
      real qdar(kdim6+maxuq),xdev(kdim6+maxuq)
      double precision dx(kdim6+maxuq),dy(kdim6+maxuq),
     +     dycal(kdim6+maxuq),dev(kdim6+maxuq),dxx(kdim6+maxuq),
     + dsumx,dsumy,dsumxx,dxn,davex,davey,davexx,dsumxxx,dsumxxxx,
     +dsumxy,dsumxxy,deno,da,db,dccc,dev2,dxex,dyex,derror(kdim6+maxuq),
     +derr(kdim6+maxuq)
c
      do 1 i=1,mmax
      do 10 j=1,nst
      if(ircode(i).eq.iabs(idata(j))) then
      dx(j)=dble(qdar(i))
      dxx(j)=dx(j)**2
      endif
   10 continue
    1 continue
      ii=0
      do 11 i=1,nst
      id=idata(i)
      if(id.lt.0) then
      ii=ii-1
      dy(i)=dble(ii-1.0)
      endif
      ii=ii+1
      dy(i)=dble(ii)
   11 continue
c      do 21 i=1,nst
c      write(*,*) i,x(i),y(i)
c   21 continue
      dsumx=0.0
      dsumy=0.0
      dsumxx=0.0
      dxn=nst
      do 2 i=1,nst
      dsumx=dsumx+dx(i)
      dsumy=dsumy+dy(i)
      dsumxx=dsumxx+dxx(i)
    2 continue
      davex=dsumx/dxn
      davey=dsumy/dxn
      davexx=dsumxx/dxn
c statement added on Nov. 24th
      dsumxx=0.0
      dsumxy=0.0
      dsumxxx=0.0
      dsumxxxx=0.0
      dsumxxy=0.0
      do 3 i=1,nst
      dx(i)=dx(i)-davex
      dxx(i)=dxx(i)-davexx
      dy(i)=dy(i)-davey
      dsumxx=dsumxx+dx(i)*dx(i)
      dsumxy=dsumxy+dx(i)*dy(i)
      dsumxxx=dsumxxx+dx(i)*dxx(i)
      dsumxxxx=dsumxxxx+dxx(i)**2
      dsumxxy=dsumxxy+dxx(i)*dy(i)
    3 continue
c      b=sumxy/sumxx
c      a=avey-b*avex
      deno=dsumxx*dsumxxxx-dsumxxx**2
      db=dsumxxxx*dsumxy-dsumxxx*dsumxxy
      db=db/deno
      dccc=dsumxx*dsumxxy-dsumxxx*dsumxy
      dccc=dccc/deno
      da=davey-db*davex-dccc*davexx
      dxex=-db/(2.0*dccc)
      dyex=da-db**2/(4.0*dccc)
      af1=sngl(da)
      bf1=sngl(db)
      cf1=sngl(dccc)
      do 4 i=1,nst
c      ycal(i)=a+b*(x(i)+avex)
c april 98 insert
      derr(i)=dsumxxxx*dx(i)**2+dsumxx*dxx(i)**2
      derr(i)=derr(i)-2.0*dsumxxx*dx(i)*dxx(i)
      derr(i)=derr(i)/deno
      derr(i)=1.0-1.0/dxn-derr(i)
      derr(i)=dsqrt(1.0/derr(i))
      dy(i)=dy(i)+davey
      dx(i)=dx(i)+davex
      dxx(i)=dxx(i)+davexx
      dycal(i)=da+db*dx(i)+dccc*dxx(i)
      if(dx(i).gt.dxex.and.dccc.lt.0.0) dycal(i)=dyex
      if(dx(i).lt.dxex.and.dccc.gt.0.0) dycal(i)=dyex
      dev(i)=dy(i)-dycal(i)
      dev(i)=dev(i)*derr(i)
    4 continue
      dev2=0.0
      do 5 i=1,nst
      dev2=dev2+dev(i)*dev(i)
      x=sngl(dx(i))
      y=sngl(dy(nst)-dy(i))
      ycal=sngl(dy(nst)-dycal(i))
      ev=sngl(-dev(i))
      idat=iabs(idata(i))
      if(ifit.eq.0) write(63,1000) i,x,y,ycal,ev,idat,
     +(ititle(idat,j),j=1,10)
      if(ifit.eq.1) write(66,1000) i,x,y,ycal,ev,idat,
     +(ititle(idat,j),j=1,10)
      if(ifit.eq.0) write(72,1000) i,x,y,ycal,ev,idat,
     +(ititle(idat,j),j=1,10)
      if(ifit.eq.1) write(75,1000) i,x,y,ycal,ev,idat,
     +(ititle(idat,j),j=1,10)
      xdev(i)=sngl(dev(i))
    5 continue
      if(ifit.eq.0) write(72,1001) nst,af1,bf1,cf1
      if(ifit.eq.1) write(75,1001) nst,af1,bf1,cf1
 1000 format(i3,4f10.5,2x,i4,2x,10a4)
 1001 format(/i3,' CALCULATED VALUES DERIVED FROM: ',3f15.8/)
      sdev=sngl(dev2/dxn)
      sdev=sqrt(sdev)
      if(icasc.eq.1) then
      do 6 i=1,mmax
        dx(i)=dble(qdar(i))
        dxx(i)=dx(i)**2
        dycal(i)=da+db*dx(i)+dccc*dx(i)**2
        dx(i)=dx(i)-davex
        dxx(i)=dxx(i)-davexx
        derror(i)=dsumxxxx*dx(i)**2+dsumxx*dxx(i)**2
        derror(i)=derror(i)-2.0*dsumxxx*dx(i)*dxx(i)
        derror(i)=dsqrt(derror(i)/deno)
        dx(i)=dx(i)+davex
        x=sngl(dx(i))
        ycal=sngl(dycal(i))
        error=sngl(derror(i))
        write(25,1000) i,x,ycal,error
c        write(72,1000) i,x,ycal,error
    6 continue
      endif
      return
      end
      SUBROUTINE FTOZ (P,ZP)
C
C ... SUBROUTINE TO COMPUTE Z FROM FREQUENCY.   EQ. 26.2.23   IN
C
C      ABRAMOWITZ, M. AND STEGUN, I.A.     (EDITORS)
C      "HANDBOOK OF MATHEMATICAL FUNCTIONS
C         WITH FORMULAS, GRAPHS, AND MATHEMATICAL TABLES"
C      PUB. BY NATIONAL BUREAU OF STANDARDS OF THE
C         U.S. DEPARTMENT OF COMMERCE,   1964.
C
C  CALLED BY SUBROUTINE NORMZ
C
      DATA  C0/2.515517/, C1/0.802853/, C2/0.010328/,
     +      D1/1.432788/, D2/0.189269/, D3/0.001308/
C
c     Modifications by FPA on 29 June 1988
      p0=p
      k=0
10    Q = P
      IF (P.GT.0.5)  Q = 1.0 - P
      TT = ALOG (1.0 / (Q * Q))
      T = SQRT (TT)
      UP = C0 + (C1 * T) + (C2 * T * T)
      DN = 1.0 + (D1 * T) + (D2 * T * T) + (D3 * T**3)
      ZP = T - (UP / DN)
      IF (P.LE.0.5)  ZP = -ZP
      if(k.eq.0) call ztof(zp,p)
      k=k+1
      if(k.eq.1) goto 10
      p=2.0 * p0 - p
      if(k.eq.2) goto 10
      RETURN
      END
      SUBROUTINE ZTOF (Z,PZ)
C
C ... SUBROUTINE TO COMPUTE FREQUENCY FROM Z.   EQ. 26.2.17   IN
C
C      ABRAMOWITZ, M. AND STEGUN, I.A.     (EDITORS)
C      "HANDBOOK OF MATHEMATICAL FUNCTIONS
C         WITH FORMULAS, GRAPHS, AND MATHEMATICAL TABLES"
C      PUB. BY NATIONAL BUREAU OF STANDARDS OF THE
C         U.S. DEPARTMENT OF COMMERCE,   1964.
C
C  CALLED BY SUBROUTINES  COMP, REORD, & WDIST
C
      DATA  PI/3.141592654/, CONST6/0.2316419/, B1/0.319381530/,
     +      B2/-0.356563782/, B3/1.781477937/, B4/-1.821255978/,
     +      B5/1.330274429/
C
      X = Z
      IF (Z.LT.0.0) X = -Z
      T = 1.0 / (CONST6 * X + 1.0)
      PID = 2.0 * PI
      XX = -X * X / 2.0
      XX = EXP (XX) / SQRT (PID)
      PZ = (B1 * T) + (B2 * T * T) + (B3 * T**3) + (B4 * T**4) +
     +     (B5 * T**5)
      PZ = 1.0 - (PZ * XX)
      IF (Z.LT.0.0) PZ = 1.0 - PZ
      RETURN
      END
      subroutine deviat(sd,af,bf,cf,unit,icasc,qdar,ifit,jvan,ivent,
     +nopt,sdopt,avesd)
c
c ... subroutine to perform event deviation analysis
c
c  added by F.P. Agterberg, August 14, 1995
c
c  called in main programme
c
      parameter(KDIM1=100,KDIM3=300,kdim4=998,KDIM6=200,MAXUQ=50)
      common n,ns,mmax,ix(kdim1,kdim3),icode(kdim4),ircode(kdim4),
     +       c(kdim6,kdim6)
      common /text/ name(kdim1,10),ititle(kdim4,10)
      integer unit,ifreq(10),jvan(kdim6,4),ifreqr(10),nopt(kdim4)
      real dev(kdim1),sd(kdim1),stats(kdim6,4),dmin(kdim6),
     + qdar(kdim6+maxuq),ycal(kdim1),xy(kdim1),af(kdim1),bf(kdim1),
     + cf(kdim1),dmax(kdim6),y(kdim1),sdopt(kdim4)
      character*1 iblank,istar,ialpha(10),ibeta(10),imora(61),irasc,
     +iminus
      character*4 iask1,iblan,isym(kdim6)
      data iask1/' *'/,iblan/'  '/
      data iblank/' '/,istar/'*'/,irasc/'M'/,iminus/'-'/
      itemp=unit
      if(ifit.eq.0) ivent=0
      if(ifit.eq.0) ivent2=0
c      do 111 i=1,ns
c      write(*,*) af(i),bf(i),cf(i)
c  111 continue
      unit=63
      if(ifit.eq.1) unit=66
      write(unit,1500)
 1500 format(///' EVENT VARIANCE ANALYSIS'/)
      xn=ns
      sumsd=0.0
      do 10 i=1,ns
      sumsd=sumsd+sd(i)**2
   10 continue
      do 20 i=1,10
      ifreq(i)=0
      ialpha(i)=iblank
      ibeta(i)=iblank
   20 continue
      if(ifit.eq.0)write(73,3050) mmax
      if(ifit.eq.1)write(76,3050) mmax
      if(ifit.eq.0) write(74,3050) mmax
      if(ifit.eq.1) write(77,3050) mmax
      if(ifit.eq.0)write(73,3051) ns
      if(ifit.eq.1) write(76,3051) ns
 3050 format(//' TOTAL NUMBER OF EVENTS = ',i3/)
 3051 format(' TOTAL NUMBER OF WELLS = ',i3)
      avesd=sqrt(sumsd/xn)
      write(unit,500) avesd
      if(ifit.eq.0) write(73,502) avesd
      if(ifit.eq.1) write(76,502) avesd
  500 format(/' AVERAGE STANDARD DEVIATION FROM LINE OF CORRELATION = ',
     +f7.4)
 5555 format('  * INDICATES EVENT WITH SD < ave SD'/)
  502 format(/' AVERAGE STANDARD DEVIATION FROM LINE OF CORRELATION = ',
     +f7.4/)
      do 101 i=1,mmax
      isum=0
      do 102 k=1,ns
      dev(k)=c(k,i)
      if(dev(k).eq.999.0) goto 102
      isum=isum+1
  102 continue
      if(ifit.eq.0) write(73,7002) ircode(i),isum,
     +(ititle(ircode(i),j),j=1,10)
      if(ifit.eq.1) write(76,7002) ircode(i),isum,
     +(ititle(ircode(i),j),j=1,10)
 7002 format(' EVENT ',i3,' OCCURS IN ',i2,' WELLS: ',10a4)
  101 continue
      do 1 i=1,mmax
      isym(i)=iblan
      write(unit,1000) ircode(i), (ititle(ircode(i),j),j=1,10)
      if(ifit.eq.0) write(73,1000)ircode(i),(ititle(ircode(i),j),j=1,10)
      if(ifit.eq.1) write(76,1000)ircode(i),(ititle(ircode(i),j),j=1,10)
 1000 format(////' EVENT NO. ',i4, 2x,10a4/)
      write(unit,2500)
      if(ifit.eq.0)write(73,2500)
      if(ifit.eq.1)write(76,2500)
 2500 format(' WELL DEVIATION     STRAT. HIGHER/LOWER THAN '
     +'EXPECTED'/)
      sum1=0.0
      sumx=0.0
      do 2 k=1,ns
c      if(k.le.kdim2/2) then
c      dev(k)=ix1(k+kdim2,i)
c      endif
c      if(k.gt.kdim2/2) then
c      dev(k)=ix(k+kdim2/2,i)
c      endif
      dev(k)=c(k,i)
c      dev(k)=dev(k)/1000.0
      if(dev(k).eq.999.0) goto 2
      sum1=sum1+1.0
      sumx=sumx+dev(k)
      if(dev(k).lt.999.0) dev(k)=dev(k)*5.0/avesd
      do 3 j=1,10
      if(dev(k).ge.0.0) idev=dev(k)+1
      if(dev(k).lt.0.0) idev=-dev(k)+1
c      idev=iabs(dev(k))+1
      if(dev(k).gt.0.0.and.j.le.idev) ialpha(j)=istar
      if(dev(k).lt.0.0.and.j.le.idev) ibeta(11-j)=istar
    3 continue
      dev(k)=dev(k)*avesd/5.0
      ycal(k)=af(k)+bf(k)*qdar(i)+cf(k)*qdar(i)**2
      peak=-bf(k)/(2.0*cf(k))
      y(k)=ycal(k)+dev(k)
      if(cf(k).gt.0.0) then
      xy(k)=bf(k)**2-4.0*cf(k)*(af(k)-y(k))
      if(xy(k).le.0.0) xy(k)=-999.0
c      write(*,*) 'CPOS',i,k,xy(k)
      if(xy(k).gt.0.0) xy(k)=(-bf(k)+sqrt(xy(k)))/(2.0*cf(k))
      if(xy(k).lt.peak.or.qdar(i).lt.peak) then
      xy(k)=-999.0
c      write(*,*) i,k,peak,qdar(i),xy(k),' CPOS'
      endif
      endif
      if(cf(k).lt.0.0) then
      xy(k)=bf(k)**2-4.0*cf(k)*(af(k)-y(k))
      if(xy(k).le.0.0) xy(k)=-999.0
c      write(*,*) 'CNEG',i,k,xy(k)
      if(xy(k).gt.0.0) xy(k)=(-bf(k)+sqrt(xy(k)))/(2.0*cf(k))
      if(xy(k).gt.peak.or.qdar(i).gt.peak) then
      xy(k)=-999.0
c      write(*,*) i,k,peak,qdar(i),xy(k),' CNEG'
      endif
      endif
      write(unit,2000) k,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
      if(ifit.eq.0) then
      write(73,2000)   k,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
      endif
      if(ifit.eq.1) then
      write(76,2000)   k,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
      endif
2000  format(i3,4x,f7.3,8x,'I',10a1,'I',10a1,'I')
      do 4 j=1,10
      ialpha(j)=iblank
      ibeta(j)=iblank
    4 continue
    2 continue
      nn=sum1
      ave=sumx/sum1
      do 222 k=1,ns
c      if(xy(k).eq.-999.0) xy(k)=ave
  222 continue
      if(dev(1).eq.999.0) xy(1)=-999.0
      dmax(i)=xy(1)
      dmin(i)=xy(1)
      if(xy(1).eq.-999.0) dmin(i)=999.0
      do 15 k=2,ns
      if(dev(k).eq.999.0) goto 15
      if(xy(k).eq.-999.0) goto 15
      if(xy(k).lt.dmin(i)) dmin(i)=xy(k)
      if(xy(k).gt.dmax(i)) dmax(i)=xy(k)
   15 continue
c      nn=sum1
c      ave=sumx/sum1
      sumxx=0.0
      sumxxx=0.0
      do 5 k=1,ns
      if(dev(k).eq.999.0) goto 5
      dev(k)=dev(k)-ave
      sumxx=sumxx+dev(k)**2
      sumxxx=sumxxx+dev(k)**3
    5 continue
      sda=sqrt(sumxx/(sum1-1.0))
      ska=sum1*sumxxx/((sum1-1.0)*(sum1-2.0)*sda**3)
      write(unit,3000) nn,ave,sda,ska
      if(ifit.eq.0)write(73,3000)   nn,ave,sda,ska
      if(ifit.eq.1)write(76,3000) nn,ave,sda,ska
      stats(i,1)=sum1
      stats(i,2)=ave
      stats(i,3)=sda
c      if(ifit.eq.0) then
      avesdh=avesd/2.0
c      if(sda.lt.avesdh) then
      if(sda.lt.avesd) then
      isym(i)=iask1
      if(ifit.eq.0) ivent=ivent+1
      endif
      tavesd=2.*avesd
      if(sda.gt.tavesd)then
      ivent2=ivent2+1
      endif
      xmmax=mmax
      crit3=.1*xmmax
      if(ivent2.gt.crit3)iwarn=iwarn+1
      if(iwarn.eq.1)then
      write(86,10000)
      write(87,10000)
      write(87,10001)
      endif
10000 format('TYPE 3 WARNING: MORE THAN 10% OF EVENTS IN OPTIMUM'/'SEQUE
     +NCE HAVE SD GREATER THAN TWICE THE AVERAGE SD')
10001 format(' ')
      stats(i,4)=-ska
 3000 format(/' SAMPLE SIZE =  ',i3,7x,'  UNCORRECTED MEAN = ',f9.3,/
     +' ADJUSTED SD = ',f7.3,'      ADJUSTED SKEWNESS = ',f8.3/)
      write(unit,3500) ave
      if(ifit.eq.0) write(73,3500) ave
      if(ifit.eq.1) write(76,3500) ave
 3500 format(/' HISTOGRAM AFTER CHANGING MEAN FROM ',f6.3,' TO 0'/)
      write(unit,4000)
      if(ifit.eq.0)write(73,4000)
      if(ifit.eq.1)write(76,4000)
 4000 format(' CLASS     LIMITS            FREQUENCY'/)
      do 6 k=1,ns
      do 7 j=1,10
cc      xk1=float(j-6)*avesd/5.0
      xk1=float(j-6)*avesd/2.5
cc      xk2=float(j-5)*avesd/5.0
      xk2=float(j-5)*avesd/2.5
      if(j.eq.1) xk1=-9.99
      if(j.eq.10) xk2=9.99
      if(dev(k).gt.xk1.and.dev(k).le.xk2) ifreq(j)=ifreq(j)+1
    7 continue
    6 continue
      do 1212 k=1,10
      ifreqr(11-k)=ifreq(k)
 1212 continue
c      do 1212 k=1,5
c      iii=ifreq(k)
c      ifreq(k)=ifreq(11-k)
c      ifreq(11-k)=iii
c 1212 continue
c      do 1213 k=1,10
c      xk1=float(k-6)*avesd/5.0
c      xk2=float(k-5)*avesd/5.0
c      if(k.eq.1) xk1=-9.99
c      if(k.eq.10)xk2=9.99
c      if(ifit.eq.0) write(73,4500) k,xk1,xk2,ifreq(k)
c      if(ifit.eq.1) write(76,4500) k,xk1,xk2,ifreq(k)
c 1213 continue
      do 12 k=1,10
cc      xk1=float(k-6)*avesd/5.0
      xk1=float(k-6)*avesd/2.5
cc      xk2=float(k-5)*avesd/5.0
      xk2=float(k-5)*avesd/2.5
      do 9 j=1,10
      if(j.le.ifreq(k)) ialpha(j)=istar
    9 continue
      if(k.eq.1) xk1=-9.99
      if(k.eq.10) xk2=9.99
      write(unit,4500) k,xk1,xk2,ifreq(k),(ialpha(j),j=1,10)
       if(ifit.eq.0) then
       write(73,4500)   k,xk1,xk2,ifreqr(k),(ialpha(j),j=1,10)
c      write(73,4500)   k,xk1,xk2,ifreq(11-k)
       endif
       if(ifit.eq.1) then
       write(76,4500)   k,xk1,xk2,ifreqr(k),(ialpha(j),j=1,10)
c      write(76,4500)   k,xk1,xk2,ifreq(11-k)
       endif
c end of change
      do 11 j=1,10
      ialpha(j)=iblank
   11 continue
      ifreq(k)=0
   12 continue
 4500 format(2x,i3,2x,f5.2,' TO ',f5.2,7x,i2,2x,'I',10a1,'I')
c make insert here
      write(unit,5503)
      write(unit,5504)
      do 223 k=1,ns
      if(dev(k).lt.999.0.and.xy(k).gt.-999.0) then
      write(unit,5502) k,dev(k),y(k),xy(k)
      endif
  223 continue
 5502 format(i3,4x,f7.3,f7.0,f9.3)
 5503 format(//'PROJECTION OF WELL DATA ONTO OPTIMUM SEQUENCE AXIS')
 5504 format(/' WELL DEVIATION LEVEL PROJECTION'/)
    1 continue
c      unit=itemp
      write(unit,4700)
 4700 format(///'SUMMARY OF EVENT VARIANCE ANALYSIS RESULTS')
c      write(unit,4800)
 4800 format('           FOR EVENT VARIANCE ANALYSIS RESULTS')
 4900 format('           SUMMARIZED IN THE FOLLOWING TABLE'//)
c      write(unit,4900)
c      write(unit,5000)
 5000 format(//45x,'N',6x,'MEAN',7x,'SD',4x,'SKEWNESS'/)
      if(icasc.eq.1) write (25,5501)
      if(icasc.eq.1) then
      write (unit,5000)
c      write(73,5000)
      endif
 5500 format(/54x,'LOWEST  AVERAGE HIGHEST'/)
 5501 format(//55x,'N',6x,'SD     LOWEST  AVERAGE HIGHEST'/)
      do 50 i=1,mmax
      istats1=stats(i,1)
      write(unit,6000) (ititle(ircode(i),j),j=1,10),
     + istats1,(stats(i,k),k=2,3),isym(i),stats(i,4)
      if(ifit.eq.0) then
      nopt(ircode(i))=istats1
      sdopt(ircode(i))=stats(i,3)
      endif
   50 continue
      write(unit,500) avesd
      write(unit,5555)
      index2=0
      write (unit,7001)
      do 1054 i=1,mmax
      istats1=stats(i,1)
      if(icasc.eq.1) write(25,6501) i,ircode(i),
     + (ititle(ircode(i),j),j=1,10),istats1,stats(i,3),
     + dmax(i),qdar(i),dmin(i),avesd
 1054 continue
 9000 write (unit,5500)
      do 54 i=1,mmax
c      istats1=stats(i,1)
c      if(icasc.eq.1) write(25,6501) i,ircode(i),
c     + (ititle(ircode(i),j),j=1,10),istats1,stats(i,3),
c     + dmax(i),qdar(i),dmin(i),avesd
      if(icasc.eq.1) write(unit,6500) i,ircode(i),
     + (ititle(ircode(i),j),j=1,10),
     + dmax(i),qdar(i),dmin(i)
      if(icasc.eq.1.and.ifit.eq.0)then
       write(74,6502) i,dmax(i),qdar(i),dmin(i),
     + ircode(i),(ititle(ircode(i),j),j=1,10)
       endif
      if(icasc.eq.1.and.ifit.eq.1)then
       write(77,6502) i,dmax(i),qdar(i),dmin(i),
     + ircode(i),(ititle(ircode(i),j),j=1,10)
       endif
   54 continue
c      rewind 75
c      do 101 i=1,mmax
c      read(75,1000) ircode(i), (ititle(ircode(i),j),j=1,10)
c      write(73,1000)ircode(i), (ititle(ircode(i),j),j=1,10)
c      read(75,2500)
c      write(73,2500)
c      do 102 k=1,ns
c      kk=k
c      read(75,2000) kk,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
c      write(73,2000)kk,dev(k),(ibeta(j),j=1,10),(ialpha(j),j=1,10)
c  102 continue
c      read(75,3000) nn,ave,sda,ska
c      write(73,3000)nn,ave,sda,ska
c      read(75,3500) ave
c      write(73,3500) ave
c      read(75,4000)
c      write(73,4000)
c      do 212 k=1,10
c      kk=k
c      read(75,4500) kk,xk1,xk2,ifreq(k),(ialpha(j),j=1,10)
c      write(73,4500)kk,xk1,xk2,ifreq(k),(ialpha(j),j=1,10)
c  212 continue
c  101 continue
c      write(unit,500) avesd
      write (unit,501)
  501 format(/'RANGE IS THE OBSERVED STRATIGRAPHIC SPREAD OF AN EVENT',/
     +'OVER ALL WELLS, PROJECTED ALONG THE OPTIMUM SEQUENCE AXIS'/)
      if(icasc.eq.1) then
      write(unit,7000)
 7000 format(//'GRAPHICAL REPRESENTATION OF EVENT RANGES: M = RASC DISTA
     +NCE'/)
 7001 format(///'ESTIMATION OF EVENT RANGES')
 7003 format(/////'STRATIGRAPHIC OVERLAP OF EVENTS (CROSS-OVER RANGE)')
c      write(unit,7001)
c 7001 format('NOTE: I IS ONLY PLOTTED IF ITS POSITION DIFFERS FROM THAT
c     +OF M'/)
 8000 format(11x,'+',14x,'+',14x,'+',14x,'+',14x,'+')
 8500 format(8x,f6.2,10x,f5.2,10x,f5.2,10x,f5.2,9x,f6.2)
      xmin=dmin(1)
      xmax=dmax(mmax)
c      if(stats(mmax,2).gt.0.0) xmax=xmax+stats(mmax,2)
      diff=xmax-xmin
      a=-diff/40.
      b=-11.*a+xmax
      x1=a+b
      x16=16.*a+b
      x31=31.*a+b
      x46=46.*a+b
      x61=61.*a+b
      write(unit,8500)x1,x16,x31,x46,x61
      write (unit,8000)
      do 51 i=1,mmax
      value1=dmin(i)
      val1=11.0+40.0*((xmax-value1)/diff)+.5
      ival1=val1
      value2=dmax(i)
      val2=11.0+40.0*((xmax-value2)/diff)+.5
      ival2=val2
      do 52 j=1,61
      imora(j)=iblank
      if(j.le.ival1.and.j.ge.ival2) imora(j)=iminus
   52 continue
c      val2=11.0+40.0*((xmax-qdar(i)-stats(i,2))/diff)+.5
c      ival2=val2
c      imora(ival2)=icor
      val3=11.0+40.0*((xmax-qdar(i))/diff)+.5
      ival3=val3
      imora(ival3)=irasc
      write(unit,7500) i,(imora(j),j=1,61),ircode(i)
     + ,(ititle(ircode(i),j),j=1,10)
   51 continue
 7500 format(1x,i3,2x,61a,7x,i3,2x,10a4)
      write(unit,8000)
      write(unit,8500)x1,x16,x31,x46,x61
      endif
      if(ifit.eq.0.and.index2.eq.0) then
      index2=1
      do 9500 i=1,mmax
      dmax(i)=jvan(i,2)-1
      dmin(i)=jvan(i,1)-1
 9500 continue
      write(unit,7003)
      goto 9000
      endif
      if(ifit.eq.1.and.index2.eq.0) then
      index2=1
      do 9501 i=1,mmax
      i4=jvan(i,4)
      i3=jvan(i,3)
      dmax(i)=qdar(i4)
      dmin(i)=qdar(i3)
 9501 continue
      write(unit,7003)
      goto 9000
      endif
      unit=itemp
 6000 format(1x,10a4,i5,2f10.3,a2,f8.3)
 6500 format(1x,i3,2x,i3,2x,10a4,2x,f7.3,1x,f7.3,1x,f7.3)
 6502 format(1x,i3,2x,f7.3,1x,f7.3,1x,f7.3,2x,i3,2x,10a4)
 6501 format(1x,i3,2x,i3,2x,10a4,i5,f10.5,1x,4(1x,f7.3))
      return
      end
