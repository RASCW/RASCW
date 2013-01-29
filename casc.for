      program casc
      real x(300),y0(13),y(300),y2(300),stats(200,2),d(300)
      real x1(200,300),y1(200,300),y12(200,300),xevent(2,200,300)
      real ypy(300),ypx(300),y100(200,300),avesd(300),ratio(300)
      real xx1(200,200),xx2(200,200),xx3(200,200),xde(200,200)
      real depthmin(300,2)
      real zz1(300),zz2(300),diff(200,200,200),diffi(20000)
      real sumdif(200,200), undif(200,200),dif(200,200,200)
      real ave(200,200),cum(200),dist(200),count(200,200)
      integer ny(300),irc(200),ievent(300),nepw(300),ievent0(300)
      integer nobs(300),ir(200),ix(200,200),ixdif(200,200,200)
      integer ic(200),irct(200),mm(200),ixi(20000),nic(200)
      character*12 inpfil0,inpfil5
      character*12 inpfil1,inpfil2,outfil,outfi1,outfi2,inpfil3,inpfil4
      character*12 outfi3,outfi4,outfi5,outfi6,outfi7,outfi8,outfi9
      character*12 outfi10,outfi11,outfi12,outfi13,outfi14,outfi15
      character*7 outfil1
      character*4 name(100,10),ititle(200,10),iititle(200,10),res1
      character*4 ititl(200,10),nititl(200,10)
      character*10 result4
      character*2 res2,res3
      character*1 code
      character*8 result,rec
      character*8 longna,longname(10)
C  DISPLAY PROGRAM NAME
      WRITE(*,'(5(/))')
      WRITE(*,'(12X,A)')
     &' CCC       A       SSS      CCC          PPPP      CCC'
      WRITE(*,'(12X,A)')
     &'C   C     A A     S   S    C   C         P   P    C   C'
      WRITE(*,'(12X,A)')
     &'C        A   A    S        C             P   P    C'
      WRITE(*,'(12X,A)')
     &'C        AAAAA     SSS     C      =====  PPPP     C'
      WRITE(*,'(12X,A)')
     &'C        A   A        S    C             P        C'
      WRITE(*,'(12X,A)')
     &'C   C    A   A    S   S    C   C         P        C   C'
      WRITE(*,'(12X,A)')
     &' CCC     A   A     SSS      CCC          P         CCC'
      WRITE (*,*)' '
      write(*,*) ' '
c      write(*,*) ' '
      WRITE (*,'(12X,A)')
     &'         CASC VERSION 3 (April,1998)'
      write (*,*) ' '
      write(*,'(12x,a)')
     &'                    by'
      write (*,*) ' '
      write(*,'(12x,a)')
     &'         F.P. Agterberg and F.M. Gradstein'
      WRITE (*,*)' '
      write (*,*)' '
c      write(*,'(5x,a)')
c     &'Enter name of dat file (*.dat): '
c      read(*,999) inpfil0
c      write(*,'(5x,a)')
c     &'Enter name of depth file (*.dep): '
c      read(*,1000) inpfil1
  999 format(a12)
 1000 format(a12)
c  998 format('     Results in File: ',a12)
c      write(*,'(5x,a)')
c     &'Enter "1" when .dep file is "decimal" ("0" otherwise): '
c      read(*,999) indec
      indec=1
c      write(*,'(5x,a)')
c     &'Enter number of lines with entries for depth: '
c      read(*,999) index
c  999 format(i2)
      index=20
c      write(*,'(5x,a)')
c     &'Enter name of RASC output file (e.g., *i.out): '
c      read(*,1000) inpfil2
      write(*,'(5x,a)')
     &'Input file names in "casctemp"; results in "*.ca1 & *.ca2"'
c      write(*,998) outfil
      open(3,file='casctemp',status='unknown')
c      open(30,file='casc-20.inc',status='unknown')
      read(3,1998) inpfil0
      read(3,1998) inpfil1
      read(3,1998) inpfil2
      read(3,1998) inpfil3
      read(3,1998) inpfil4
      read(3,1998) inpfil5
      open(4,file=inpfil0)
      open(5,file=inpfil1)
      open(25,file=inpfil2)
      open(26,file=inpfil3)
      open(27,file=inpfil4)
      open(30,file=inpfil5)
c      open(6,file='outfil',status='unknown')
c      open(7,file='cascwinR',status='unknown')
c      open(8,file='cascwinS',status='unknown')
      call date_and_time(result)
      res1=result(1:4)
      res2=result(5:6)
      res3=result(7:8)
      result4=res1//'/'//res2//'/'//res3
  777 if(inpfil2(12:12).eq.' ')then
      inpfil2=' '//inpfil2
      goto 777
      endif
      outfil1=inpfil2
      outfil=outfil1//'j.out'
      outfi1=outfil1//'.ca1'
      outfi2=outfil1//'.ca2'
      outfi3=outfil1//'.sd1'
      outfi4=outfil1//'.sd2'
      outfi5=outfil1//'.de1'
      outfi6=outfil1//'.de2'
      outfi7=outfil1//'.de3'
      outfi8=outfil1//'.de4'
      outfi9=outfil1//'.ds1'
      outfi10=outfil1//'.dem'
      outfi11=outfil1//'.df1'
      outfi12=outfil1//'.di1'
      outfi13=outfil1//'.dn2'
      outfi14=outfil1//'.di2'
      outfi15=outfil1//'.df2'
      do 10102 i=1,50
      if(outfil(1:1).eq.' ') outfil=outfil(2:50)//' '
      if(outfi1(1:1).eq.' ') outfi1=outfi1(2:50)//' '
      if(outfi2(1:1).eq.' ') outfi2=outfi2(2:50)//' '
      if(outfi3(1:1).eq.' ') outfi3=outfi3(2:50)//' '
      if(outfi4(1:1).eq.' ') outfi4=outfi4(2:50)//' '
      if(outfi5(1:1).eq.' ') outfi5=outfi5(2:50)//' '
      if(outfi6(1:1).eq.' ') outfi6=outfi6(2:50)//' '
      if(outfi7(1:1).eq.' ') outfi7=outfi7(2:50)//' '
      if(outfi8(1:1).eq.' ') outfi8=outfi8(2:50)//' '
      if(outfi9(1:1).eq.' ') outfi9=outfi9(2:50)//' '
      if(outfi10(1:1).eq.' ') outfi10=outfi10(2:50)//' '
      if(outfi11(1:1).eq.' ') outfi11=outfi11(2:50)//' '
      if(outfi12(1:1).eq.' ') outfi12=outfi12(2:50)//' '
      if(outfi13(1:1).eq.' ') outfi13=outfi13(2:50)//' '
      if(outfi14(1:1).eq.' ') outfi14=outfi14(2:50)//' '
      if(outfi15(1:1).eq.' ') outfi15=outfi15(2:50)//' '
10102 continue
      open(6,file=outfil,status='unknown')
      open(7,file=outfi1,status='unknown')
      open(8,file=outfi2,status='unknown')
      open(9,file=outfi3,status='unknown')
      open(10,file=outfi4,status='unknown')
      open(11,file=outfi7,status='unknown')
      open(12,file=outfi8,status='unknown')
      open(13,file=outfi5,status='unknown')
      open(14,file=outfi6,status='unknown')
      open(15,file=outfi9,status='unknown')
      open(16,file=outfi10,status='unknown')
      open(17,file=outfi11,status='unknown')
      open(18,file=outfi12,status='unknown')
      open(19,file=outfi13,status='unknown')
      open(20,file=outfi14,status='unknown')
      open(21,file=outfi15,status='unknown')
      write(6,905) inpfil2,outfil,result4
  905 format(1x,'CASC RESULTS FOR ',a12,18x,'File: ',a12,2x,a10/)
      read(30,9000)kcrito,decrito,icut,imean,isqrt,iendo,kcrit,irascs,io
      ialpha=0
      call cascpage
      read(26,1200)
      read(26,1200)
      read(27,1200)
      read(27,1200)
      read(26,7000) iscale,nsapril
      read(27,7000) iscale,nsapril
 7000 format(i2,23x,i3)
      read(26,1200)
      if(iscale.gt.0) read(27,1200)
      write(9,10071) nsapril,ift
      write(11,10071) nsapril,ift
      if(iscale.gt.0) then
      write(10,10071) nsapril,ift
      write(12,10071) nsapril,ift
      endif
10071 format(//,' TOTAL NUMBER OF WELLS = ',2i3/)
      do 8001 i=1,nsapril
      read(26,7001) isapril,nedapril
      if(iscale.gt.0) read(27,1200)
 7001 format(8x,i2,6x,i3)
 7002 format(2x,10a4)
      nobs(i)=nedapril
      write(9,50601) isapril,nedapril
      write(11,50601) isapril,nedapril
      if(iscale.gt.0) then
      write(10,50601)isapril,nedapril
      write(12,50601)isapril,nedapril
      endif
50601 format(' WELL # ',i2,' WITH ',i3,' EVENTS')
 8001 continue
 1998 format(a12)
      ik=0
      kmax=0
c      if(indec.gt.0) read(5,1200)
      read(5,1999) rec
 1999 format(a8)
      if(rec(1:3).ne.'DEC') then
      indec=0
      rewind 5
      endif
      do 5 i=1,100
      read(4,2000)(name(i,j),j=1,10)
      if(name(i,1).eq.'LAST'.or.name(i,1).eq.'last') goto 7
      jj=0
    4 read(4,5000)(ievent0(j),j=1,20)
      do 6 j=1,20
c      nepw(i)=j+jj-1
      if(ievent0(j).gt.0) jj=jj+1
      if(ievent0(j).eq.-999) then
      nepw(i)=jj
      goto 5
      endif
    6 continue
c      jj=jj+20
      goto 4
    5 continue
    7 rewind 4
c change from 50 to 300 in next statement
      do 1 i=1,300
    3 read (5,2000) (name(i,j),j=1,10)
      if(name(i,1).eq.'    '.or.name(i,1).eq.'0000') goto 3
      if(name(i,1).eq.'LAST'.or.name(i,1).eq.'last') goto 30
      kmax=kmax+1
      read (5,3000) code,rth
      if(indec.gt.0) read(5,1200)
      k=0
      ny(i)=0
      nd=13
      if(indec.gt.0) nd=9
      do 40 ii=1,index
      if(ny(i).ne.0) goto 1
      if(indec.eq.0) read (5,4000) (y0(j),j=1,nd)
      if(indec.gt.0) read (5,4001) (y0(j),j=1,nd)
c      if(ny(i).ne.0) goto 40
      do 20 j=1,nd
      if(y0(j).eq.0.0) then
      ny(i)=j+k-1
      goto 40
      endif
      y0(j)=y0(j)-rth
      if(code.eq.'f'.or.code.eq.'F') y0(j)=0.30480*y0(j)
      y(j+k)=y0(j)
      x(j+k)=j+k
      x1(i,j+k)=x(j+k)
      y1(i,j+k)=y(j+k)
   20 continue
      k=k+nd
      if(k.eq.nepw(i)) goto 1
   40 continue
 2000 format(10a4)
 3000 format(a1,f5.0)
 4000 format(13f6.0)
 4001 format(9(f7.0,1x))
    1 continue
c30    do 22 i=1,kmax
c      do 21 j=1,ny(i)
c      write(*,*) x1(i,j),y1(i,j)
c   21 continue
c   22 continue
30    do 1001 i=1,kmax
      read (4,2000) (name(i,j),j=1,10)
      jj=0
 1040 read (4,5000) (ievent0(j),j=1,20)
 5000 format(20i4)
      do 1020 j=1,20
      nepw(i)=j+jj-1
      if(ievent0(j).eq.-999) goto 1111
      ievent(j+jj)=ievent0(j)
      xevent(1,i,j+jj)=ievent(j+jj)
      if(ievent(j+jj).lt.0) xevent(1,i,j+jj)=-xevent(1,i,j+jj)
 1020 continue
      jj=jj+20
      goto 1040
 1111 ki=1
      do 10 k=1,nepw(i)
      if(ievent(k).gt.0) then
      xevent(2,i,k)=y1(i,ki)
      if(ievent(k+1).gt.0) ki=ki+1
      endif
      if(ievent(k).lt.0) then
      xevent(2,i,k)=y1(i,ki)
      if(ievent(k+1).gt.0) ki=ki+1
      endif
   10 continue
 1001 continue
c      do 22 i=1,kmax
c      do 21 j=1,nepw(i)
c      write(*,*) xevent(1,i,j),xevent(2,i,j)
c   21 continue
c   22 continue
      read(25,1200)
      read(25,1200)
      read(25,1200)
      read(25,1200)
      read(25,1200)
      write(6,1140)
 1140 format(1x,'OPTIMUM SEQUENCE RESULTS')
 1150 format(///1x,'SCALED OPTIMUM SEQUENCE RESULTS')
  111 ik=ik+1
      if(ik.gt.1) write(6,1150)
      if(ik.gt.1) read(25,1200)
      do 100 k=1,kmax
      do 99 kk=1,4
      read(25,1200)
   99 continue
      read(25,1250) ned
      xned=ned
      read(25,1200)
      read(25,1200)
      read(25,1200)
      read(25,1200)
 1200 format(1x)
 1250 format(16x,i3)
 1100 format(i3,3f10.5)
      do 101 ii=1,200
      read(25,1100) iii,xopt,ycal,error
      if(iii.eq.0) goto 100
      xde(k,ii)=xopt
      xx1(k,ii)=ycal
      xx2(k,ii)=sqrt(1.0/xned+error**2)
      xx3(k,ii)=0.0
      aaa=1.0-1.0/xned-error**2
      if(aaa.gt.0.0) xx3(k,ii)=sqrt(aaa)
  101 continue
  100 continue
      read(25,1200)
      read(25,1200)
      read(25,1200)
      ift=0
      mmax=0
      do 102 ii=1,200
      read(25,6000)iii,irc(ii),ir(ii),(ititle(ii,j),j=1,10),stats(ii,1),
     + stats(ii,2),avesd(ii)
      if(irascs.eq.0.and.ialpha.eq.0) then
      nic(ii)=irc(ii)
      do 71500 j=1,10
      nititl(ii,j)=ititle(ii,j)
71500 continue
      endif
      irct(ii)=irc(ii)
      mm(ii)=iii
      ic(ii)=irc(ii)
      do 715 j=1,10
      ititl(ii,j)=ititle(ii,j)
  715 continue
      if(iii.eq.0.and.irc(ii).eq.0) goto 103
      if(iii.eq.0.and.irc(ii).eq.1) then
      ialpha=1
      if(ir(ii).gt.0) ift=1
      goto 103
      endif
      mmax=mmax+1
  102 continue
 6000 format(1x,i3,2x,i3,i2,10a4,f5.0,f10.5,26x,f7.3)
  103 continue
      if(ir(ii).gt.0) ift=1
c      write(*,*) mmax
      do 401 i=1,kmax
      if(ik.eq.1) then
      do 8010 ii=1,7
      read(26,1200)
      if(iscale.gt.0) read(27,1200)
 8010 continue
      write(9,5060) (name(i,j),j=1,10)
      if(iscale.gt.0) write(10,5060)(name(i,j),j=1,10)
 5060 format(///2x,10a4/)
 5058 format(2x,'i',4x,'X(i)',4x,'YMAX-Y(i)',1x,'EXPECTED',2x,
     +'DEVIATION',4x,'DEPTH',1x,'  NO.  NAME'/)
      write(9,5058)
      if(iscale.gt.0) write(10,5058)
15058 format(2x,'i     X     OBSERVED CALCULATED  NO.  NAME'/)
      endif
      nep=nepw(i)
      jj=0
      do 402 k=1,mmax
      do 403 j=1,nep
      ievent(j)=xevent(1,i,j)
      if(irc(k).eq.ievent(j)) then
      jj=jj+1
      y100(i,jj)=xevent(2,i,j)
      endif
  403 continue
  402 continue
      njj=jj
c      write(*,*) njj
      if(ik.eq.1) then
      do 8012 j=1,njj
c      if(ik.eq.1) then
      read(26,8000)ia,xa,ya,ycala,eva,idata,
     +(iititle(idata,jj),jj=1,10)
 8000 format(i3,4f10.5,2x,i4,2x,10a4)
      write(9,801)ia,xa,ya,ycala,eva,xevent(2,i,j),idata,
     +(iititle(idata,jj),jj=1,10)
 801  format(i3,4f10.5,f10.1,2x,i4,2x,10a4)
c      endif
      if(iscale.gt.0) then
      read(27,8000)ia,xa,ya,ycala,eva,idata,
     +(iititle(idata,jj),jj=1,10)
      write(10,801)ia,xa,ya,ycala,eva,xevent(2,i,j),idata,
     +(iititle(idata,jj),jj=1,10)
      endif
 8012 continue
      if(i.lt.kmax) then
      do 8011 k=1,3
      read(26,1200)
      write(9,1200)
      if(iscale.gt.0) then
      read(27,1200)
      write(10,1200)
      endif
 8011 continue
      endif
      endif
      nex=njj+1
      do 23 k=1,njj-1
      do 24 j=k+1,njj
      if(y100(i,j).lt.y100(i,k)) then
      y100(i,nex)=y100(i,k)
      y100(i,k)=y100(i,j)
      y100(i,j)=y100(i,nex)
      endif
   24 continue
   23 continue
      jj=1
      x1(i,1)=1.
      y1(i,1)=y100(i,1)
      do 404 j=1,njj
      if(j.gt.1.and.y100(i,j).ne.y100(i,j-1)) then
      jj=jj+1
      x1(i,jj)=jj
      y1(i,jj)=y100(i,j)
      endif
  404 continue
      ny(i)=jj
c      if(i.eq.1) then
c      do 22 k=1,njj
c      write(*,*) k,y100(i,k)
c   22 continue
c      do 21 j=1,ny(i)
c      write(*,*) j,x1(i,j),y1(i,j)
c   21 continue
c      endif
  401 continue
      do 200 i=1,kmax
      n=ny(i)
      do 201 j=1,n
c added statements
      if(j.eq.1) then
      depthmin(i,1)=y1(i,j)
      if(ift.gt.0) depthmin(i,1)=depthmin(i,1)/3.0480
      endif
      if(j.eq.2) then
      depthmin(i,2)=y1(i,j)
      if(ift.gt.0) depthmin(i,2)=depthmin(i,2)/3.0480
      endif
      x(j)=x1(i,j)
      y(j)=y1(i,j)
  201 continue
      ypy(i)=(y(2)-y(1))/(x(2)-x(1))
      ypx(i)=(y(n)-y(n-1))/(x(n)-x(n-1))
      yp1=ypy(i)
      ypn=ypx(i)
      call spline(x,y,n,yp1,ypn,y2)
      do 202 j=1,n
      y12(i,j)=y2(j)
  202 continue
  200 continue
      mmax=mmax
      if(ik.eq.1) write(7,3010) kmax,avesd(1)
      if(ik.gt.1) write(8,3010) kmax,avesd(1)
      if(ik.eq.1) write(7,30010)mmax,ift
      if(ik.gt.1) write(8,30010)mmax,ift
 3010 format(/////'TOTAL NUMBER OF WELLS = ',i3,6x,'AVE SD = 'f7.3/)
30010 format('TOTAL NUMBER OF EVENTS = ',2i3/)
c      do 500 i=1,kmax
c      write(7,3001)(name(i,kkk),kkk=1,10), nepw(i)
c  500 continue
      do 300 i=1,kmax
      write(6,2004) i,(name(i,kkk),kkk=1,10)
      if(ik.eq.1) write(7,2004) i,(name(i,kkk),kkk=1,10)
      if(ik.gt.1) write(8,2004) i,(name(i,kkk),kkk=1,10)
      if(ik.eq.1) write(11,5060)(name(i,j),j=1,10)
      if(ik.gt.1.and.iscale.gt.0) write(12,5060)(name(i,j),j=1,10)
      write(6,2001)
      write(6,2002)
      if(ik.eq.1)write(7,2011)
      if(ik.gt.1)write(8,2011)
      if(ik.eq.1)write(7,2012)
      if(ik.gt.1)write(8,2012)
      if(ik.eq.1) write(11,15058)
      if(ik.gt.1.and.iscale.gt.0) write(12,15058)
 2001 format(45x,'PROB   MIN   MAX   OBS   MIN   MAX')
 2011 format(/'PROB   MIN   MAX   OBS  RATIO  MIN   MAX')
 2012 format('DEPTH DEPTH DEPTH DEPTH        OBS   OBS'/)
 2002 format(45x,'DEPTH DEPTH DEPTH DEPTH  OBS   OBS'/)
 2003 format(1x,i3,1x,10a4,f5.0,1x,f5.0,1x,f5.0,1x,f5.0,1x,f5.0,1x,f5.0)
 2005 format(4(f5.0,1x),f5.3,1x,2(f5.0,1x),i3,1x,10a4)
 2006 format(f5.0,1x,f5.0,1x,f5.0,'  * * *',18x,i3,1x,10a4)
 2007 format(' * * *',12x,f5.0,1x,f5.3,'  * *',8x,i3,1x,10a4)
 2008 format(' * * *')
 2004 format(///i3,1x,10a4)
 3001 format(10a4,' CONTAINS ',i3,' EVENTS')
      n=ny(i)
      do 301 j=1,n
      x(j)=x1(i,j)
      y(j)=y1(i,j)
      y2(j)=y12(i,j)
      if(ift.gt.0)then
      y(j)=y(j)/3.0480
      y2(j)=y2(j)/3.0480
      endif
  301 continue
      kk=0
      kkk=0
      do 302 k=1,mmax
      ratio(k)=stats(k,2)/avesd(1)
      xx=xx1(i,k)
      call splint(x,y,n,y2,xx,yy)
      yy1=yy
      if(yy1.lt.0.0) yy1=0.0
      xx=xx-2.0*stats(k,2)*xx2(i,k)
      call splint(x,y,n,y2,xx,yy)
      yy2=yy
      if(yy2.lt.0.0) yy2=0.0
      xx=xx+4.0*stats(k,2)*xx2(i,k)
      call splint(x,y,n,y2,xx,yy)
      yy3=yy
      if(yy3.lt.0.0) yy3=0.0
      xx=xx1(i,k)
      xx=xx-2.0*stats(k,2)*xx3(i,k)
      call splint(x,y,n,y2,xx,yy)
      yy4=yy
      if(yy4.lt.0.0) yy4=0.0
c      if(yy4.lt.0.0) yy4=20.0
      xx=xx+4.0*stats(k,2)*xx3(i,k)
      call splint(x,y,n,y2,xx,yy)
      yy5=yy
      if(yy5.lt.0.0) yy5=0.0
      d(k)=0.0
      nep=nepw(i)
      do 303 j=1,nep
      ievent(j)=xevent(1,i,j)
      if(irc(k).eq.ievent(j)) then 
      d(k)=xevent(2,i,j)
      if(ift.gt.0) d(k)=d(k)/3.0480
      endif
  303 continue
c      kpal=0
c      do 305 k9=1,mmax
c      if(d(k9).gt.0.) then 
c      kpal=kpal+1
c      if(kpal.eq.1) depthmin=d(k9)
c      if(d(k9).lt.depthmin) depthmin=d(k9)
c      endif
c  305 continue
      d200=200.
      if(ift.gt.0) d200=.3048*d200
      if(k.gt.1.and.xx1(i,k).gt.xx1(i,k-1)) then
      if(xx2(i,k).le.1.0.and.yy2.gt.(yy1-d200)) then
      if(yy3.lt.(yy1+d200).and.yy1.gt.yy2.and.yy1.lt.yy3) then
      if(d(k).eq.0.0) then
      write(6,2003)irc(k),(ititle(k,jjj),jjj=1,10),yy1,yy2,yy3
      if(ik.eq.1)then
      write(7,2006)yy1,yy2,yy3,irc(k),(ititle(k,jjj),jjj=1,10)
      endif
      if(ik.gt.1)then
      write(8,2006)yy1,yy2,yy3,irc(k),(ititle(k,jjj),jjj=1,10)
      endif
      goto 302
c      yy4=0.0
c      yy5=0.0
      endif
      if(d(k).ne.0.0)then
      write(6,2003)irc(k),(ititle(k,jjj),jjj=1,10),yy1,yy2,yy3,d(k),yy4,
     +yy5
1101  format(i3,3f10.4,i4,2x,10a4)
      if(ik.eq.1) then
      kk=kk+1
      write(7,2005)yy1,yy2,yy3,d(k),ratio(k),yy4,yy5,irc(k)
     +,(ititle(k,jjj),jjj=1,10)
      write(11,1101)kk,xde(i,k),d(k),yy1,irc(k),(ititle(k,jk),jk=1,10)
      endif
      if(ik.gt.1) then
      kkk=kkk+1
      write(8,2005)yy1,yy2,yy3,d(k),ratio(k),yy4,yy5,irc(k)
     +,(ititle(k,jjj),jjj=1,10)
      write(12,1101)kkk,xde(i,k),d(k),yy1,irc(k),(ititle(k,jk),jk=1,10)
      endif
      goto 302
      endif
      endif
      endif
      endif
      if(d(k).gt.0.0) then
      if(ik.eq.1) then
      kk=kk+1
      if(kk.eq.1) then
      yy1=depthmin(i,1)
      ddd=d(k)
      endif
      if(kk.eq.2) yy1=depthmin(i,1)
      if(kk.eq.2.and.d(k).gt.ddd) yy1=depthmin(i,2)
      if(kk.eq.nobs(i)) yy1=d(k)
      write(11,1101)kk,xde(i,k),d(k),yy1,irc(k),(ititle(k,jk),jk=1,10)
      write(7,2007) d(k),ratio(k),irc(k),(ititle(k,jjj)
     +,jjj=1,10)
      endif
      if(ik.gt.1) then
      kkk=kkk+1
      if(kkk.eq.1) then
      yy1=depthmin(i,1)
      ddd=d(k)
      endif
      if(kkk.eq.2) yy1=depthmin(i,1)
      if(kkk.eq.2.and.d(k).gt.ddd) yy1=depthmin(i,2)
      if(kkk.eq.nobs(i)) yy1=d(k)
      write(12,1101)kkk,xde(i,k),d(k),yy1,irc(k),(ititle(k,jk),jk=1,10)
      write(8,2007) d(k),ratio(k),irc(k),(ititle(k,jjj)
     +,jjj=1,10)
      endif
      endif
      if(d(k).eq.0.0) then
      if(ik.eq.1) write(7,2008)
      if(ik.gt.1) write(8,2008)
      endif
  302 continue
  300 continue
      if(ik.eq.1.and.ialpha.eq.1) goto 111
      rewind 11
      rewind 12
      do 399 i=1,111
      read(11,398)longname
      write(13,398)longname
      longna=longname(1)
      if(longna(3:3).eq.'i') goto 397
      if(ialpha.eq.1) then
      read(12,398)longname
      write(14,398)longname
      endif
  399 continue
  397 continue
  398 format(10a8)
      read(11,398) longname
      write(13,398) longname
      if(ialpha.eq.1) then
      read(12,398) longname
      write(14,398) longname
      read(12,398) longname
      write(14,398) longname
      endif
      do 600 i=1,kmax
      do 601 j=1,nobs(i)
      read(11,1101)m,xde(i,j),d(j),zz1(j),irc(j),(ititle(j,jk),jk=1,10)
  601 continue
      do 602 j=1,nobs(i)-1
      do 603 k=j+1,nobs(i)
      if(zz1(j).gt.zz1(k)) then
      zz=zz1(k)
      zz1(k)=zz1(j)
      zz1(j)=zz
      endif
  603 continue
  602 continue
      do 6010 j=1,nobs(i)
      d10=d(j)
      z10=zz1(j)
      if(ift.gt.0)then
      d10=10.*d10
      z10=10.*z10
      endif
      id10=d10
      d10=id10
ccc if added on March 31, 2007
c      if(z10.lt.0.0) z10=0.0
      write(13,1101)j,xde(i,j),d10,z10,irc(j),(ititle(j,jk),jk=1,10)
 6010 continue
      if(ialpha.eq.1) then
      do 6011 j=1,nobs(i)
      read(12,1101)m,xde(i,j),d(j),zz2(j),irc(j),(ititle(j,jk),jk=1,10)
 6011 continue
      do 604 j=1,nobs(i)-1
      do 605 k=j+1,nobs(i)
      if(zz2(j).gt.zz2(k)) then
      zz=zz2(k)
      zz2(k)=zz2(j)
      zz2(j)=zz
      endif
  605 continue
  604 continue
      do 6012 j=1,nobs(i)
c      xde(i,j)=xde(i,j)-1.
      d10=d(j)
      z20=zz2(j)
ccc statement inserted on April 3rd, 2007
      if(j.eq.nobs(i).and.z20.gt.2.0*zz2(j-1)) z20=2.0*zz2(j-1) 
      if(ift.gt.0)then
      d10=10.*d10
      z20=10.*z20
      endif
      id10=d10
      d10=id10
c      if(z20.lt.0.0) z20=0.0
      write(14,1101)j,xde(i,j),d10,z20,irc(j),(ititle(j,jk),jk=1,10)
 6012 continue
      endif
      if(i.lt.kmax) then
      do 606 ii=1,7
      read(11,398)longname
      write(13,398)longname
      if(ialpha.eq.1) then
      read(12,398)longname
      write(14,398)longname
      endif
  606 continue
      endif
  600 continue

ccc     TESTING casc-20
      write(15,7051)
 7051 format(1x,'DEPTH SCALING RESULTS'/)
c      read(30,9000)kcrito,decrito,icut,imean,isqrt,iendo,kcrit,irascs,io
 9000 format(i2,f8.2,7i2)
      iend=0
      imean=1
      icut=1
      iam=3
      if(kcrit.eq.0) kcrit=1
      if(kcrit.gt.1) iendo=1
      if(io.eq.0) io=1
      kkmax=kcrit
      if(io.gt.kcrit) kkmax=io
      xio=io
      rewind 13
      do 696 i=1,111
      read(13,398)longname
      longna=longname(1)
      if(longna(3:3).eq.'i') goto 697
  696 continue
  697 continue
      read(13,398) longname
      do 699 i=1,kmax
      do 698 k=1,kcrit
      count(i,k)=0.0
      sumdif(i,k)=0.0
  698 continue
  699 continue
ccc added on Dec.30
      do 722 i=1,mmax
      do 723 j=1,nobs(i)
      do 724 k=1,kkmax
      dif(i,j,k)=0.001
      diff(i,j,k) = 0.001
  724 continue
  723 continue
  722 continue
ccc
      sum=0.0
      ntotal=0
      do 800 i=1,kmax
      if(i.gt.1.and.i.le.kmax) then
      do 795 ii=1,7
      read (13,398) longname
c      write (15,398) longname
  795 continue
      endif
      do 8801 j=1,nobs(i)
      read(13,1101)m,xde(i,j),d(j),zz1(j),irc(j),(ititle(j,jk),jk=1,10)
      ix(i,j)=int(xde(i,j))+1
c      write(15,1102)m,ix(i,j),d(j),zz1(j),irc(j),(ititle(j,jk),jk=1,10)
c 1102 format(i3,i6,4x,2f10.4,i4,2x,10a4)
 8801 continue
      sum=sum+d(nobs(i))-d(1)
      ntotal=ntotal+nobs(i)
  800 continue
      ntotal=ntotal-i
      total=ntotal+1.0
      avedd=sum/total
      bound1=-(xio*avedd)**0.5
      bound2=(xio*avedd)**0.5
c      decrit=decrito+avedd
c      decritn=-decrito+avedd
c      if(imean.eq.0) then
c      decrit=decrito
c      decritn=-decrito
c      endif
      write(15,8020) sum,total,avedd
 8020 format (3f12.5)
      rewind 13
      rewind 14
      do 796 i=1,111
      if(irascs.eq.0) read(13,398)longname
      if(irascs.eq.1) read(14,398)longname
      write(15,398) longname
      longna=longname(1)
      if(longna(3:3).eq.'i') goto 797
  796 continue
  797 continue
      if(irascs.eq.0) read(13,398) longname
      if(irascs.eq.1) read(14,398) longname
      write(15,398) longname
ccc
      do 700 i=1,kmax
      if(i.gt.1.and.i.le.kmax) then
      do 695 ii=1,7
      if(irascs.eq.0) read (13,398) longname
      if(irascs.eq.1) read (14,398) longname
      write (15,398) longname
  695 continue
      endif
      if(irascs.eq.0) then
      do 701 j=1,nobs(i)
      read(13,1101)m,xde(i,j),d(j),zz1(j),irct(j),(ititle(j,jk),jk=1,10)
      ix(i,j)=int(xde(i,j))+1
      write(15,1102)m,ix(i,j),d(j),zz1(j),irct(j),(ititle(j,jk),jk=1,10)
 1102 format(i3,i6,4x,2f10.4,i4,2x,10a4)
  701 continue
      endif
      if(irascs.eq.1) then
      do 1701 j=1,nobs(i)
      read(14,1101)m,xde(i,j),d(j),zz1(j),irc(j),(ititle(j,jk),jk=1,10)
      nds=irc(j)
      do 2797 k=1,200
      if(nds.eq.irct(k)) then
      ix(i,j)=mm(k)
      goto 1702
      endif
 2797 continue
 1702 write(15,1102)m,ix(i,j),d(j),zz1(j),irc(j),(ititle(j,jk),jk=1,10)
 1701 continue
      endif
      do 702 j=1,nobs(i)
      do 725 k=1,kkmax
      if(j.le.(nobs(i)-k)) then
      differ=d(j+k)-d(j)
      diff(i,j,k)=differ
      if(isqrt.eq.1) then
      xk=k
      diff(i,j,k)=(abs(differ-xk*avedd))**0.5
      endif
      if(differ.lt.xk*avedd.and.isqrt.eq.1) diff(i,j,k)=-diff(i,j,k)
      endif
  725 continue
      if(kcrito.lt.kcrit) kcrito=kcrit
      do 705 k=1,kcrito
      xk=k
      if(j.le.(nobs(i)-k)) then
      difference=d(j+k)-d(j)
      if(isqrt.eq.0) then
      decrit=decrito+xk*avedd
      decritn=-decrito+xk*avedd
      if(difference.gt.decrit.or.difference.lt.decritn) goto 705
      endif
c      if(difference.gt.decrit) difference=decrit
c      if(difference.lt.decritn) difference=decritn
      dif(i,j,k)=difference
      if(isqrt.eq.1) then
      if(difference.ge.0.0) dif(i,j,k)=dif(i,j,k)**0.5
      if(difference.lt.0.0) dif(i,j,k)=-(abs(difference))**0.5
      endif
      ixdif(i,j,k)=ix(i,j+k)-ix(i,j)
      if(ixdif(i,j,k).gt.(kcrito+k-1)) goto 702
 7049 format(5i3,3f10.3,i4)
      count(ix(i,j),ix(i,j+k))=count(ix(i,j),ix(i,j+k))+1.0
      sumdif(ix(i,j),ix(i,j+k))=sumdif(ix(i,j),ix(i,j+k))+dif(i,j,k)
      if(j.lt.5) write(15,7049) i,j,k,ix(i,j),ix(i,j+k),dif(i,j,k),count
     +(ix(i,j),ix(i,j+k)),sumdif(ix(i,j),ix(i,j+k)),ixdif(i,j,k)
      endif                            
  705 continue
 7048 format(2i4,10f12.3)
  702 continue
  700 continue
      do 720 i=1,mmax
      do 721 j=1,nobs(i)
      if(isqrt.eq.0) write(17,7048) i,j,(diff(i,j,kk),kk=1,kkmax)
      if(isqrt.eq.1) write(21,7048) i,j,(diff(i,j,kk),kk=1,kkmax)
  721 continue
  720 continue

cc sorting of diff(i,j,1) inserted on 01/09/06
      k=0
      do 1720 i=1,mmax
      do 1721 j=1,nobs(i)
      if(diff(i,j,io).ne.0.001) then
      k=k+1
      diffi(k)=diff(i,j,io)
c      write(17,17048) k,diffi(k)
      endif
 1721 continue
 1720 continue
17048 format(i4,2f12.3)
      kend=k
      do 1711 i=1,kend
      ixi(i)=i
 1711 continue
      do 1712 i=1,kend-1
      do 1713 j=i+1,kend
      if(diffi(i).le.diffi(j)) goto 1713
      temp2=diffi(i)
      diffi(i)=diffi(j)
      diffi(j)=temp2
      temp1=ixi(i)
      ixi(i)=ixi(j)
      ixi(j)=temp1
 1713 continue
 1712 continue
      zzend=kend
      sxy=0.
      sxx=0.
      do 1714 i=1,kend-1
      zzi=i
      zzi=zzi/zzend
      call ftoz(zzi,zi)
      if(isqrt.eq.0) write(17,17048) i,zi,diffi(i)
      if(isqrt.eq.1) write(21,17048) i,zi,diffi(i)
      if(diffi(i).lt.bound1.or.diffi(i).gt.bound2) then
      sxy=sxy+zi*diffi(i)
      sxx=sxx+diffi(i)*diffi(i)
      endif
 1714 continue
      b=sxy/sxx
      if(isqrt.ne.0) write(21,17049) b,io
17049 format(//'Slope of best-fit =',f10.5,';  Order =',i2) 
cc

      do 703 i=1,mmax-1
      do 704 k=i+1,i+kcrito
      undif(i,k)=0.0
      if(count(i,k).gt.0.0) undif(i,k)=sumdif(i,k)/count(i,k)
      write(15,7050) i,k,count(i,k),sumdif(i,k),undif(i,k)
  704 continue
  703 continue
 7050 format(2i3,f5.0,2f10.3)
      if(irascs.eq.0) then
      do 70700 i=1,mmax
      ic(i)=nic(i)
      do 70701 jk=1,10
      ititl(i,jk)=nititl(i,jk)
70701 continue
70700 continue
      endif
      do 706 i=1,mmax-1
      ave(i,1)=undif(i,i+1)
      if(count(i,i+1).lt.iam) ave(i,1)=0.0
      xcrit=0.0
      weight10=0.0
      weight20=0.0
      ave10=0.0
      weight1=0.0
      weight2=0.0
      ave1=0.0
      ave2=0.0
      do 707 k=1,kcrit-1
      if(i.ge.kcrit.and.i.le.(mmax-kcrit)) then
      if(count(i,i+k+1).ge.iam.and.count(i+1,i+k+1).ge.iam) then
      weight10=count(i,i+k+1)*count(i+1,i+k+1)
      weight1=weight10/(count(i,i+k+1)+count(i+1,i+k+1))
      ave1=undif(i,i+k+1)-undif(i+1,i+k+1)
      endif
      if(count(i,i+k+1).lt.iam.or.count(i+1,i+k+1).lt.iam) weight1=0.0
      if(count(i-k,i+1).ge.iam.and.count(i-k,i).ge.iam) then
      weight20=count(i-k,i+1)*count(i-k,i)
      weight2=weight20/(count(i-k,i+1)+count(i-k,i))
      ave2=undif(i-k,i+1)-undif(i-k,i)
      endif
      if(count(i-k,i+1).lt.iam.or.count(i-k,i).lt.iam) weight2=0.0
      ave10=ave10+ave1*weight1+ave2*weight2
      xcrit=xcrit+weight1+weight2
      endif
      if(i.eq.27)then
      write(15,90000)i,k,count(i,i+k+1),count(i+1,i+k+1),
     +count(i-k,i+1),count(i-k,i)
      write(15,90000)i,k,undif(i,i+k+1),undif(i+1,i+k+1),
     +undif(i-k,i+1),undif(i-k,i)
      write (15,90000)i,k,weight10,weight1,ave1,we
     +ight20,weight2,ave2,xcrit,ave10
      endif
90000 format(2i2,8f10.3)
  707 continue
      if(count(i,i+1).ge.iam) xcrit=xcrit+count(i,i+1)
      ave(i,kcrit)=ave10+ave(i,1)*count(i,i+1)
      if(xcrit.gt.0.0) ave(i,kcrit)=ave(i,kcrit)/xcrit
      if(i.lt.kcrit.or.i.gt.(mmax-kcrit)) ave(i,kcrit)=ave(i,1)
      write(15,7052) i,count(i,i+1),ave(i,1),ave(i,kcrit),ic(i),(ititl(i
     +,jk),jk=1,10)
  706 continue
 7052 format(i3,3f10.3,2x,i3,2x,10a4)
      write(15,7056)
 7056 format(//)
  726 if(iend.eq.1) then
      do 727 i=1,mmax-1
      ave(i,1)=ave(i,kcrit)
  727 continue
      endif
      cum(1)=0.0
      do 709 i=2,mmax
      cum(i)= cum(i-1)+ave(i-1,1)
  709 continue
      do 710 i=1,mmax
      write(15,7055) ic(i),cum(i),(ititl(i,jk),jk=1,10)
  710 continue
 7055 format(i4,f10.4,3x,10a4)
      write(15,7056)
      do 711 i=1,mmax
      ixi(i)=i
  711 continue
      do 712 i=1,mmax-1
      do 713 j=i+1,mmax
      if(cum(i).le.cum(j)) goto 713
      temp2=cum(i)
      cum(i)=cum(j)
      cum(j)=temp2
      temp1=ixi(i)
      ixi(i)=ixi(j)
      ixi(j)=temp1
  713 continue
  712 continue
      cum1=cum(1)
      if(cum1.lt.0.0) then
      do 7810 i=1,mmax
      cum(i)=cum(i)-cum1
 7810 continue
      endif
      do 714 i=1,mmax
      write(15,7055) ic(ixi(i)),cum(i),(ititl(ixi(i),jk),jk=1,10)
      if(iend.eq.0) write(18,7059) i,ic(ixi(i)),cum(i)
      if(iend.eq.1) write(20,7059) i,ic(ixi(i)),cum(i)
 7059 format(2i4,f10.4)
  714 continue
      write(15,7056)
      do 716 i=1,mmax-1
      dist(i)=cum(i+1)-cum(i)
  716 continue
      write(15,7056)
      dist(mmax)=-9.9999
      do 717 i=1,mmax
      write(15,7055) ic(ixi(i)),dist(i),(ititl(ixi(i),jk),jk=1,10)
  717 continue
      if(iend.eq.0) write(16,7053)mmax
      if(iend.eq.1) write(19,7053)mmax
 7053 format('***'/,1x,i10)
      xmin=dist(1)
      xmax=xmin
      do 708 i=1,mmax-1
      if(dist(i).lt.xmin) xmin=dist(i)
      if(dist(i).gt.xmax) xmax=dist(i)
  708 continue
      dx=(xmax-xmin)/25.0
      xmin=xmin-dx
      xmax=xmax+dx
      if(iend.eq.0) write(16,7054) xmax,xmin
      if(iend.eq.1) write(19,7054) xmax,xmin
 7054 format(1x,2f17.8)
      do 718 i=1,mmax-1
      if(iend.eq.0) write(16,7057) ic(ixi(i)),dist(i),(ititl(ixi(i),jk),
     +jk=1,10)
      if(iend.eq.1) write(19,7057) ic(ixi(i)),dist(i),(ititl(ixi(i),jk),
     +jk=1,10)
  718 continue
      if(iend.eq.0) write(16,7058) ic(ixi(mmax)),dist(mmax),(ititl(ixi(m
     +max),jk),jk=1,10)
      if(iend.eq.1) write(19,7058) ic(ixi(mmax)),dist(mmax),(ititl(ixi(m
     +max),jk),jk=1,10)
 7057 format(i4,f10.4,3x,10a4,',    0.1')
 7058 format(i4,f10.4,3x,10a4,',')
      if(iend.eq.0.and.iendo.gt.0) write(15,7061)
 7061 format(/,' RESULTS OF NEW RUN',/)
      iend=iend+1
      if (iend.eq.1.and.kcrit.gt.1.and.iendo.gt.0) goto 726
      write(16,7162) isqrt,irascs,kcrit
      write(17,7162) isqrt,irascs,kcrit
      write(18,7162) isqrt,irascs,kcrit
      write(19,7162) isqrt,irascs,kcrit
      write(20,7162) isqrt,irascs,kcrit
      write(21,7162) isqrt,irascs,kcrit
 7162 format(" Graph parameters: ",3i2)
      stop
      end
      subroutine spline(x,y,n,yp1,ypn,y2)
      real x(300),y(300),y2(300),u(300)
      y2(1)=0.
      u(1)=0.
c      y2(1)=-0.5
c      u(1)=(3./(x(2)-x(1)))*((y(2)-y(1))/(x(2)-x(1))-yp1)
c      write(*,*) u(1)
      do 1 i=2,n-1
         sig=(x(i)-x(i-1))/(x(i+1)-x(i-1))
         p=sig*y2(i-1)+2.
         y2(i)=(sig-1.)/p
         u(i)=(6.*((y(i+1)-y(i))/(x(i+1)-x(i))-(y(i)-y(i-1))
     *        /(x(i)-x(i-1)))/(x(i+1)-x(i-1))-sig*u(i-1))/p
c      write(*,*) i,u(i)
    1 continue
      qn=0.
      un=0.
c      qn=0.5
c      un=(3./(x(n)-x(n-1)))*(ypn-(y(n)-y(n-1))/(x(n)-x(n-1)))
      y2(n)=(un-qn*u(n-1))/(qn*y2(n-1)+1.)
      do 2 k=n-1,1,-1
         y2(k)=y2(k)*y2(k+1)+u(k)
    2 continue
      return
      end
      subroutine splint(x,y,n,y2,xx,yy)
      real x(300),y(300),y2(300)
      klo=1
      khi=n
    1 if(khi-klo.gt.1) then
         k=(khi+klo)/2
         if(x(k).gt.xx) then
            khi=k
         else
            klo=k
         endif
      goto 1
      endif
      xh=x(khi)-x(klo)
      if(xh.eq.0.) pause 'bad x input in splint'
      a=(x(khi)-xx)/xh
      b=(xx-x(klo))/xh
      yy=a*y(klo)+b*y(khi)+
     *         ((a**3-a)*y2(klo)+(b**3-b)*y2(khi))*(xh**2)/6.
      return
      end
      SUBROUTINE CASCPAGE
C
      WRITE(6,'(3(/))')
      WRITE(6,'(6X,A)')
     +'          CCC      A      SSS     CCC         PPPP     CCC'
      WRITE(6,'(6X,A)')
     +'         C   C    A A    S   S   C   C        P   P   C   C'
      WRITE(6,'(6X,A)')
     +'         C       A   A   S       C            P   P   C'
      WRITE(6,'(6X,A)')
     +'         C       AAAAA    SSS    C      ===== PPPP    C'
      WRITE(6,'(6X,A)')
     +'         C       A   A       S   C            P       C'
      WRITE(6,'(6X,A)')
     +'         C   C   A   A   S   S   C   C        P       C   C'
      WRITE(6,'(6X,A)')
     +'          CCC    A   A    SSS     CCC         P        CCC'
      WRITE(6,'(2(/))')
      WRITE(6,'(6X,A)')
     +'                 CASC Version 3 (April, 1998)'
      write(6,'(1(/))')
      write(6,'(6x,a)')
     +'                            by'
      write(6,'(1(/))')
      write(6,'(6x,a)')
     +'                 F.P. Agterberg and F.M. Gradstein'
      write(6,'(1(/))')
      write(6,'(6x,a)')
     +'     CORRELATION AND STANDARD-ERROR CALCULATION OF FOSSIL EVENTS'
      write(6,'(6x,a)')
     +'     ___________________________________________________________'
      WRITE(6,'(5(/))')
      WRITE(6,'(5X,A)')
     +'CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC'//
     +'CCCCCCCCCCCCCCCCCCC'
      WRITE(6,'(/)')
      RETURN
      END
      subroutine ftoz(p,zp)
      data  c0/2.515517/, c1/0.802853/, c2/0.010328/,
     +      d1/1.432788/, d2/0.189269/, d3/0.001308/
      p0=p
      k=0
   10 q=p
      if(p.gt.0.5) q=1.0-p
      tt=alog(1.0/(q*q))
      t=sqrt(tt)
      up=c0+(c1*t)+(c2*t*t)
      dn=1.0+(d1*t)+(d2*t*t)+(d3*t**3)
      zp=t-(up/dn)
      if(p.le.0.5) zp=-zp
      if(k.eq.0) call ztof(zp,p)
      k=k+1
      if(k.eq.1) goto 10
      p=2.0*p0-p
      if(k.eq.2) goto 10
      return
      end
      subroutine ztof(z,pz)
      data  pi/3.141592654/, const6/0.2316419/, b1/0.319381530/,
     +      b2/-0.356563782/, b3/1.781477937/, b4/-1.821255978/,
     +      b5/1.330274429/
      x=z
      if(z.lt.0.0) x=-z
      t=1.0/(const6*x+1.0)
      pid=2.0*pi
      xx=-x*x/2.0
      xx=exp(xx)/sqrt(pid)
      pz=(b1*t)+(b2*t*t)+(b3*t**3)+(b4*t**4)+(b5*t**5)
      pz=1.0-(pz*xx)
      if(z.lt.0.0) pz=1.0-pz
      return
      end




