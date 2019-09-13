#-*-coding:euc-kr-*-

import sys
import os
import re
import xlrd
import xlsxwriter
from calendar import monthrange


# state description
#
#     0 ---> 1 -----> 2
#        |       ^
#        |->-----| 
#
# 0 normal
# 1 first entry after 4
# 2 first entry after 6
#
def get_info( inFile ):
    wb= xlrd.open_workbook( inFile )
    ws= wb.sheet_by_index( 0 )

    # config
    NAME= ws.row_values( 1,5,6 )[0]
    x= re.search( '(\d+)-(\d+)-(\d+)\s+(\d+):(\d+):(\d+)', ws.row_values( 1,1,2 )[0] )
    YEAR= int( x.group( 1 ) )
    MONTH= int( x.group( 2 ) )
    DAYS= monthrange( YEAR, MONTH )[1]

    return [NAME, YEAR, MONTH, DAYS]


def get_timestamp( date ):
    x= re.search( '(\d+)-(\d+)-(\d+)\s+(\d+):(\d+):(\d+)', date )
    return [int(x.group(1)),int(x.group(2)),int(x.group(3)),int(x.group(4)),int(x.group(5)),int(x.group(6))]


def extract_db_old( REF_TIME, inFile ):
    # open the worksheet
    wb= xlrd.open_workbook( inFile )
    ws= wb.sheet_by_index( 0 )

    # config
    #REF_TIME= 5

    # FSM variables
    st = 0
    month_change= 0

    # FSM starts
    DB=[]
    for ridx in range( 2,ws.nrows ):
        pL= get_timestamp( ws.row_values( ridx-1,1,2 )[0] )
        cL= get_timestamp( ws.row_values( ridx,1,2 )[0]   )

        if   st == 0: # get start_time
            if ( cL[3] >= REF_TIME ):
                DB.append( cL ) # start_time of today
                st = 1
        elif st == 1:
            if   ( cL[2]  > pL[2] ):
                DB.append( pL ) # end_time candidate of today
                st= 2
            elif ( cL[1]  > pL[1] ):
                DB.append( pL )
                st= 3
        elif st == 2:
            if   ( pL[3]  < REF_TIME ) and ( cL[3] >= REF_TIME ):
                DB.pop(  )
                DB.append( pL ) # overtime work. end_time of today
                DB.append( cL ) # start_time of tomorrow
                st= 1
            elif ( pL[3] >= REF_TIME ) and ( cL[3] >= REF_TIME ):
                DB.append( pL ) # start_time of tomorrow
                st= 1
        elif st==3:
            if   ( pL[3]  < REF_TIME ) and ( cL[3] >= REF_TIME ):
                DB.pop( )
                DB.append( pL )
                break
            elif ( pL[3] >= REF_TIME ) and ( cL[3] >= REF_TIME ):
                break

    return DB

def extract_db( REF_TIME, inFile ):
    # open the worksheet
    wb= xlrd.open_workbook( inFile )
    ws= wb.sheet_by_index( 0 )

    # config
    #REF_TIME= 5

    # FSM variables
    st = 0
    month_change= 0

    # FSM starts
    DB=[]
    for ridx in range( 2,ws.nrows ):
        pL= get_timestamp( ws.row_values( ridx-1,1,2 )[0] )
        cL= get_timestamp( ws.row_values( ridx,1,2 )[0]   )

        if   st == 0: # get start_time
            if ( cL[3] >= REF_TIME ):
                DB.append( cL ) # start_time of today
                st = 1
        elif st == 1:
            if   ( cL[2] != pL[2] ):
                DB.append( pL ) # end_time candidate of today
                month_change = 1 if cL[1] > pL[1] else 0
                st= 2
        elif st == 2:
            if   ( pL[3]  < REF_TIME ) and ( cL[3] >= REF_TIME ):
                DB.pop(  )
                DB.append( pL ) # overtime work. end_time of today
                if month_change == 1: break
                DB.append( cL ) # start_time of tomorrow
                st= 1
            elif ( pL[3] >= REF_TIME ) and ( cL[3] >= REF_TIME ):
                if month_change == 1: break
                DB.append( pL ) # start_time of tomorrow
                st= 1

    return DB


def get_work_time( st,et ):
    h= 24+et[3]-st[3] if st[2] != et[2] else et[3]-st[3]
    m= et[4]-st[4]
    s= et[5]-st[5]
    if s < 0:
        s= s+60
        m= m-1
    if m < 0:
        m= m+60
        h= h-1

    wt= h + m/60 + s/3600

    if   wt >= 22.5:
        h= h-2
        m= m-30
    elif wt >= 18.0:
        h= h-2
    elif wt >= 13.5:
        h= h-1
        m= m-30
    elif wt >=  9.0:
        h= h-1
    elif wt >=  4.5:
        m= m-30
    if m<0:
        m= m+60
        h= h-1

    return [ h,m,s ]


def get_sum_worktime( DB ):
    sh= 0
    sm= 0
    ss= 0
    for db in DB:
        sh += db[9][0]
        sm += db[9][1]
        ss += db[9][2]

    ( ssq,ssr ) = ( int( ss/60 ), ss % 60 )
    ss  = ssr
    sm += ssq
    ( smq,smr ) = ( int( sm/60 ), sm % 60 )
    sm  = smr
    sh += smq
    return [sh,sm,ss]

def get_week_average( DAYS, wt ):
    sum_in_sec= wt[0]*3600 + wt[1]*60 + wt[2]
    wavg = int(((( sum_in_sec * 7 ) / DAYS)*10+10-1)//10)
    h = int( wavg/3600 )
    m = int( ( wavg - h * 3600 )/60 )
    s = wavg - h*3600 - m*60
    return [h,m,s]

def generate_summary( iFN, DB ):
    # debug
    DBG_STDOUT=0
    DBG_GENLOG=0

    info= get_info( iFN )
    NAME=  info[0]
    YEAR=  info[1]
    MONTH= info[2]
    DAYS=  info[3]

    FDB=[]
    for day in range( DAYS ):
        FDB.append( [ YEAR, MONTH, day+1, 0,0,0, 0,0,0,[0,0,0] ] )

    for i in range( 0, len( DB ), 2 ):
        FDB[DB[i][2]-1]= [   DB[i][0],   DB[i][1],   DB[i][2],
                        DB[i][3],   DB[i][4],   DB[i][5],
                        DB[i+1][3], DB[i+1][4], DB[i+1][5],
                        get_work_time( DB[i], DB[i+1] ) ]

    #for fdb in FDB:
        #print( fdb )

    wt= get_sum_worktime( FDB )
    week_avg = get_week_average( DAYS, wt )

    #print('DAYS:',DAYS)

    wb= xlsxwriter.Workbook( NAME+'.xlsx' )
    ws_stat=  wb.add_worksheet( 'Stat' )
    ws_stat.set_column( 0,3,12 )
    ws_stat.write( 0,0,'이름' )
    ws_stat.write( 0,1,NAME )
    ws_stat.write( 1,0,'기간' )
    ws_stat.write( 1,1,'%4d-%02d' % ( YEAR,MONTH ) )
    ws_stat.write( 2,0,'합계' )
    ws_stat.write( 2,1,'%d:%02d:%02d' % ( wt[0],wt[1],wt[2] ) )
    ws_stat.write( 3,0,'주평균' )
    ws_stat.write( 3,1,'%d:%02d:%02d' % ( week_avg[0],week_avg[1],week_avg[2] ) )

    ws_stat.write( 5,0,'날짜' )
    ws_stat.write( 5,1,'최초출입' )
    ws_stat.write( 5,2,'최후출입' )
    ws_stat.write( 5,3,'근무시간' )

    for r,fdb in enumerate( FDB ):
        ws_stat.write( r+6,0, '%4d-%02d-%02d' % ( fdb[0],fdb[1],fdb[2] ) )
        if (fdb[3]==0) and (fdb[4]==0) and (fdb[5]==0) and (fdb[6]==0) and (fdb[7]==0) and (fdb[8]==0):
            ws_stat.write( r+6,1, '-' )
            ws_stat.write( r+6,2, '-' )
            ws_stat.write( r+6,3, '-' )
        else:
            ws_stat.write( r+6,1, '%2d:%02d:%02d' % ( fdb[3],fdb[4],fdb[5] ) )
            ws_stat.write( r+6,2, '%2d:%02d:%02d' % ( fdb[6],fdb[7],fdb[8] ) )
            ws_stat.write( r+6,3, '%2d:%02d:%02d' % ( fdb[9][0],fdb[9][1],fdb[9][2] ) )
    wb.close( )

    if DBG_STDOUT==1:
        oFN= os.path.splitext( os.path.abspath( iFN ) )[0]+'.log'
        print('NAME:',NAME, 'YEAR:',YEAR, 'MONTH:',MONTH, 'SUM:', wt[0], wt[1], wt[2], 'WAVG:', week_avg[0], week_avg[1], week_avg[2] )
        with open( oFN, 'wt' ) as fo:
            print( '%10s   %8s     %8s %8s' % ( 'date','start','end', 'work' ) )
            print( '-------------------------------------------' )
            for i,edb in enumerate(DB):
                if i % 2 == 0: print( '%4d-%02d-%02d | %02d:%02d:%02d --> ' % ( edb[0], edb[1], edb[2], edb[3], edb[4], edb[5] ), end='' )
                else         : print( '%02d:%02d:%02d %s' % ( edb[3], edb[4], edb[5], get_work_time( DB[i-1], DB[i] ) ) )

    if DBG_GENLOG==1:
        oFN= os.path.splitext( os.path.abspath( iFN ) )[0]+'.log'
        with open( oFN, 'wt' ) as fo:
            print( '%10s   %8s     %8s %8s' % ( 'date','start','end', 'work' ) ,file=fo)
            print( '-------------------------------------------' ,file=fo)
            for i,edb in enumerate(DB):
                if i % 2 == 0: print( '%4d-%02d-%02d | %02d:%02d:%02d --> ' % ( edb[0], edb[1], edb[2], edb[3], edb[4], edb[5] ), end='' ,file=fo)
                else         : print( '%02d:%02d:%02d %s' % ( edb[3], edb[4], edb[5], get_work_time( DB[i-1], DB[i] ) ) ,file=fo)



if __name__ == '__main__':

    iFN= sys.argv[1]
    eDB= extract_db( 5, iFN )
    generate_summary( iFN, eDB )
