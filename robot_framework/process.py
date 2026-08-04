from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from datetime import datetime, timedelta
import pyodbc
import requests
from requests_ntlm import HttpNtlmAuth
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

def tjek_case(cur):
    cases_expired = {}
    cur.execute("""
        SELECT aktid,
            Udleveringsmappelink,
            Dokumentlistemappelink,
            sharepoint_udleveringslink,
            last_run_complete,
            handler_email
        FROM dbo.cases
        WHERE slettet_sharepoint = 0
        OR slettet_go = 0
        AND status = 'Afsluttet'
    """)

    rows = cur.fetchall()
    columns = [col[0] for col in cur.description]
    rows = [dict(zip(columns, row)) for row in rows]

    cutoff = (datetime.now() - timedelta(days=31)).date()

    for element in rows:
        last_run = element["last_run_complete"]
        if last_run is None:
            continue
        last_run = datetime.strptime(str(last_run)[:10], "%Y-%m-%d").date()

        if last_run and last_run < cutoff:
            aktid = element["aktid"]

            cases_expired[aktid] = {
                "Udleveringsmappelink": element["Udleveringsmappelink"],
                "Dokumentlistemappelink": element["Dokumentlistemappelink"],
                "Sharepointmappelink": element["sharepoint_udleveringslink"],
                "aktid": aktid,
                "handler_email": element['handler_email']
            }

    return cases_expired

def anonymiser_sag(cur, aktid):
    """
    Anonymiserer en sag baseret på aktid.
    Berørte kolonner:
      - dbo.cases: citizen_name, citizen_id, citizen_email, Beskrivelse
      - dbo.case_journal_items: subject, body
    """

    # Anonymiser dbo.cases
    cur.execute("""
        UPDATE dbo.cases
        SET
            citizen_name  = 'ANONYMISERET',
            citizen_id    = 'ANONYMISERET',
            citizen_email = 'ANONYMISERET',
            Beskrivelse   = 'ANONYMISERET',
            status = 'ANONYMISERET'
        WHERE aktid = ?
    """, (aktid,))

    cases_affected = cur.rowcount

    # Anonymiser dbo.case_journal_items via case_aktid (FK direkte til aktid)
    cur.execute("""
        UPDATE dbo.case_journal_items
        SET
            subject = 'ANONYMISERET',
            body    = 'ANONYMISERET',
            payload = 'ANONYMISERET',
            to_email = 'ANONYMISERET'
            
        WHERE case_aktid = ?
    """, (aktid,))

    journal_affected = cur.rowcount

    return {
        "aktid": aktid,
        "cases_opdateret": cases_affected,
        "journal_items_opdateret": journal_affected,
    }

def delete_case_go(go_api_url, session, sagsnummer, cur, aktid):
    '''
    Deletes case in go
    '''
    try:
        url = f"{go_api_url}/aktindsigter/_goapi/Cases/{sagsnummer}"
        response = session.delete(url, data= {"Data": ""}, timeout=1200)
        response.raise_for_status()
        marker_slettet_go(cur, aktid)
        return True
    except:
        orchestrator_connection.log_error('Failed in deleting go case')
        return False

def marker_slettet_sharepoint(cur, aktid):
    """
    Markerer en sag som slettet i Sharepoint.
    """
    cur.execute("""
        UPDATE dbo.cases
        SET
            slettet_sharepoint = 1
        WHERE aktid = ?
    """, (aktid,))

    affected = cur.rowcount

    return {
        "aktid": aktid,
        "rækker_opdateret": affected,
    }

def marker_slettet_go(cur, aktid):
    """
    Markerer en sag som slettet i GO.
    """
    cur.execute("""
        UPDATE dbo.cases
        SET
            slettet_go         = 1
        WHERE aktid = ?
    """, (aktid,))

    affected = cur.rowcount

    return {
        "aktid": aktid,
        "rækker_opdateret": affected,
    }

def sharepoint_client(username: str, password: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))
    # Authenticate to SharePoint using Office365 credentials

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")

    cert_credentials = {
        "tenant": api.username,
        "client_id": api.password,
        "thumbprint": certification.username,
        "cert_path": certification.password
    }

    ctx = ClientContext(sharepoint_site_url).with_client_certificate(**cert_credentials)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    orchestrator_connection.log_info(f"✅ Authenticated to SharePoint. Site Title: {web.properties['Title']}")
    return ctx

def delete_sharepoint_folder(folder_path: str, ctx: ClientContext, orchestrator_connection: OrchestratorConnection, aktid, cursor):
    """
    Recursively deletes a SharePoint folder and all its contents.
    """
    orchestrator_connection.log_info(f"🗑 Deleting folder: {folder_path}")
    try:
        target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        ctx.load(target_folder)
        ctx.execute_query()

        files = target_folder.files
        ctx.load(files)
        ctx.execute_query()
        for file in files:
            orchestrator_connection.log_info(f"  Deleting file: {file.serverRelativeUrl}")
            file.delete_object()
        ctx.execute_query()

        subfolders = target_folder.folders
        ctx.load(subfolders)
        ctx.execute_query()
        for subfolder in subfolders:
            delete_sharepoint_folder(subfolder.serverRelativeUrl, ctx, orchestrator_connection, aktid, cursor)

        target_folder.delete_object()
        ctx.execute_query()
        orchestrator_connection.log_info(f"✅ Folder deleted: {folder_path}")
        marker_slettet_sharepoint(cursor, aktid)
        return True
    except Exception as e:
        orchestrator_connection.log_info(f'An exception occurred: {e}')
        return False
    
def create_ntlm_session(username: str, password: str) -> requests.Session:
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    return session

def process(orchestrator_connection: OrchestratorConnection) -> None:
    sql_server = orchestrator_connection.get_constant("SqlServer").value  
    conn_string = f"DRIVER={{SQL Server}};SERVER={sql_server};DATABASE=AKTINDSIGTERPERSONALEMAPPER;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    conn.autocommit = False
    cur = conn.cursor()
    
    expired_cases = tjek_case(cur)

    if expired_cases:
        #Go stuff
        goapiurl = orchestrator_connection.get_constant('GOApiURL').value
        go_api = orchestrator_connection.get_credential('GOAktApiUser')
        go_user = go_api.username
        go_password = go_api.password


        #Sharepointstuff
        sharepointapi = orchestrator_connection.get_credential('SharePointAPI')
        sharepointuser = sharepointapi.username
        sharepointpassword = sharepointapi.password
        sharepointsiteurl = orchestrator_connection.get_constant('PersonaleIndsigtSharepointUrl').value
        ctx = sharepoint_client(username = sharepointuser, password = sharepointpassword, sharepoint_site_url= sharepointsiteurl, orchestrator_connection= orchestrator_connection)
    
        for case in expired_cases.values():
            #sharepoint deleter
            folder_path_dokumentliste = case['Dokumentlistemappelink'].rsplit('.com')[-1]
            aktid = case['aktid']
            orchestrator_connection.log_info(f'Deleting {aktid}')
            try:
                delete_sharepoint_folder(folder_path = folder_path_dokumentliste, ctx = ctx, orchestrator_connection= orchestrator_connection, aktid = aktid, cursor = cur)
                orchestrator_connection.log_info(f'Deleted sharepoint folder for case {case["aktid"]}')
            except:
                orchestrator_connection.log_info(f'Delete error in dokumentliste sharepoint for case {aktid}')

            try:
                folder_path_udlevering = case['Sharepointmappelink'].rsplit('.com')[-1]
                delete_sharepoint_folder(folder_path = folder_path_udlevering, ctx = ctx, orchestrator_connection= orchestrator_connection, aktid = aktid, cursor = cur)
            except:
                orchestrator_connection.log_info(f'Delete error in udlevering sharepoint for case {aktid}')

            #Go deleter
            session = create_ntlm_session(username = go_user, password= go_password)
            udleveringsmappe_id = case["Udleveringsmappelink"].rsplit("/")[-1]
            try:
                delete_case_go(go_api_url= goapiurl, session = session, sagsnummer = udleveringsmappe_id, cur = cur, aktid= aktid)
                orchestrator_connection.log_info(f'Deleted go folder for case {case["aktid"]}')
            except:
                orchestrator_connection.log_info(f'Delete error in go for {aktid}')

            #Sql stuff
            anonymiser_sag(cur = cur, aktid= aktid)
            conn.commit()

    else:
        print('Nothing to clean today')
