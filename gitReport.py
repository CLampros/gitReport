from textwrap import wrap
import time
import json
import time
import calendar
import os

from github import Github
import xlsxwriter
from dotenv import load_dotenv


def get_xlsx_obj():
    # Create a workbook and add a worksheet.
    timestr = time.strftime("%Y%m%d-%H%M%S")
    workbook = xlsxwriter.Workbook(f'gitReport-{timestr}.xlsx')

    return workbook


def save_and_close_xlsx_obj(workbook):
    # Save and close xlsx.
    workbook.close()


def sleep_if_core_rate_limit_reached(g):
    core_rate_limit_remaining = g.get_rate_limit().core.remaining
    core_rate_limit = g.get_rate_limit().core
    reset_timestamp = calendar.timegm(core_rate_limit.reset.timetuple())
    sleep_time = reset_timestamp - calendar.timegm(time.gmtime()) + 5
    if core_rate_limit_remaining < 10:
        print(f"Core rate limit remaining ==> {core_rate_limit_remaining}")
        print(f"Sleeping for {sleep_time}...")
        print()
        time.sleep(sleep_time)


def get_orgs_list(g):
    orgslist = [orgs.login for orgs in g.get_user().get_orgs()]
        
    return orgslist


def get_orgs_repo_details(g, o, xlsx):
    # Create a worksheet.
    worksheet = xlsx.add_worksheet('UUC REPO BRANCHES')
    # Formats.
    fwrap = xlsx.add_format({
        'align': 'center',
        'text_wrap': True,
        'valign': 'vcenter'
        })

    fcenter = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })

    fheaders = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'blue',
        'font_color': 'white',
        'font_size': 14
    })

    fnon_compliant = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_color': 'red',
    })

    rows = {
        "repo": 1,
        "private": 2,
        "branches": 3,
        "has_prod": 4,
        "protected": 5,
        "has_teams": 6,
        "author_name": 7,
        "author_email": 8,
    }
    worksheet.set_column('B:B', 50, fcenter)
    worksheet.set_column('C:C', 10, fwrap)
    worksheet.set_column('D:D', 30, fwrap)
    worksheet.set_column('E:E', 15, fwrap)
    worksheet.set_column('F:F', 15, fwrap)
    worksheet.set_column('G:G', 15, fwrap)
    worksheet.set_column('H:H', 30, fwrap)
    worksheet.set_column('I:I', 50, fwrap)
    
    # Headers.
    worksheet.write(1, rows["repo"], "Repo", fheaders)
    worksheet.write(1, rows["private"], "Private", fheaders)
    worksheet.write(1, rows["branches"], "Branches", fheaders)
    worksheet.write(1, rows["has_prod"], "Has Prod", fheaders)
    worksheet.write(1, rows["protected"], "Protected", fheaders)
    worksheet.write(1, rows["has_teams"], "Has Teams", fheaders)
    worksheet.write(1, rows["author_name"], "Author name", fheaders)
    worksheet.write(1, rows["author_email"], "Author e-mail", fheaders)

    row = 2
    for org in g.get_user().get_orgs():
        if org.login == o:
            for repo in org.get_repos():
                # Name. If it is not compliant is getting updated with red font color later on.
                worksheet.write(row, rows["repo"], repo.name)
                # Is private?
                worksheet.write(row, rows["private"], repo.private)
                # Branches.
                branches = [branch.name for branch in repo.get_branches()]
                branches_str = "\n".join(branches)
                worksheet.write(row, rows["branches"], branches_str)
                # Has prod?
                has_prod = 'prod' in branches
                worksheet.write(row, rows["has_prod"], has_prod)

                # Has teams?
                teams = repo.get_teams()

                if teams:
                    worksheet.write(row, rows["has_teams"], "TRUE")
                else:
                    worksheet.write(row, rows["has_teams"], "FALSE")
                    worksheet.write(row, rows["repo"], repo.name, fnon_compliant)

                # Last commit on prod branch (if exists).
                if has_prod:
                    # Is protected?
                    protected = repo.get_branch('prod').protected
                    worksheet.write(row, rows["protected"], protected)
                    if not protected:
                        worksheet.write(row, rows["repo"], repo.name, fnon_compliant)
                        
                    # Commiter_name & commiter_email.
                    try:
                        commiter_name = repo.get_branch('prod').commit.author.name
                        commiter_email = repo.get_branch('prod').commit.author.email
                        worksheet.write(row, rows["author_name"], commiter_name)
                        worksheet.write(row, rows["author_email"], commiter_email)
                    except AttributeError:
                        worksheet.write(row, rows["author_name"], 'No commits')
                        worksheet.write(row, rows["author_email"], 'No commits')
                else:
                    worksheet.write(row, rows["repo"], repo.name, fnon_compliant)
                    worksheet.write(row, rows["protected"], 'No prod')
                    worksheet.write(row, rows["author_name"], 'No prod')
                    worksheet.write(row, rows["author_email"], 'No prod')
                row += 1


def get_teams_repo_permissions(g, o, xlsx):
    # Create a worksheet.
    worksheet = xlsx.add_worksheet('Teams Permissions')
    # Formats.
    fwrap = xlsx.add_format({
        'text_wrap': True,
        'valign': 'vcenter'
        })

    fcenter = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })

    fheaders = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'blue',
        'font_color': 'white',
        'font_size': 14
    })

    rows = {
        "team": 1,
        "repo": 2,
        "permissions": 3,
    }
    worksheet.set_column('B:B', 30, fcenter)
    worksheet.set_column('C:C', 30, fcenter)
    worksheet.set_column('D:D', 20, fwrap)
    
    # Headers.
    worksheet.write(1, rows["team"], "Team", fheaders)
    worksheet.write(1, rows["repo"], "Repo", fheaders)
    worksheet.write(1, rows["permissions"], "Permissions", fheaders)

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for team in org.get_teams():
                for repo in team.get_repos():
                    # Team.
                    worksheet.write(row, rows["team"], team.name)

                    # Repo.
                    worksheet.write(row, rows["repo"], repo.name)

                    # Permissions.
                    json_permissions = json.dumps(team.get_repo_permission(repo).raw_data, indent = 4)
                    json_permissions = json_permissions.replace('{', '')
                    json_permissions = json_permissions.replace('}', '')
                    worksheet.write(row, rows["permissions"], json_permissions)

                    row += 1


def get_memebers_repo_permissions(g, o, xlsx):
    # Create a worksheet.
    worksheet = xlsx.add_worksheet('Members Permissions')
    # Formats.
    fwrap = xlsx.add_format({
        'text_wrap': True,
        'valign': 'vcenter'
        })

    fcenter = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })

    fheaders = xlsx.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'blue',
        'font_color': 'white',
        'font_size': 14
    })

    worksheet.set_column('B:B', 30, fcenter)
    worksheet.set_column('C:C', 30, fcenter)
    worksheet.set_column('D:D', 15, fcenter)
    worksheet.set_column('E:E', 25, fwrap)

    rows = {
        "repo": 1,
        "collaborators": 2,
        "repo_admin": 3,
        "permissions": 4,
    }
    
    # Headers.
    worksheet.write(1, rows["repo"], "Repo", fheaders)
    worksheet.write(1, rows["collaborators"], "Collaborator", fheaders)
    worksheet.write(1, rows["repo_admin"], "Repo admin", fheaders)
    worksheet.write(1, rows["permissions"], "Permissions", fheaders)

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for repo in org.get_repos():
                worksheet.write(row, rows["repo"], repo.name)
                for collaborator in repo.get_collaborators():
                    # Sleep if no more requests are available until they have been replenished.
                    sleep_if_core_rate_limit_reached(g)
                    # Permissions.
                    # https://pygithub.readthedocs.io/en/latest/github_objects/Repository.html
                    # https://docs.github.com/en/rest/reference/collaborators#get-repository-permissions-for-a-user
                    # https://docs.github.com/en/enterprise-server@3.2/organizations/managing-access-to-your-organizations-repositories/viewing-people-with-access-to-your-repository
                    json_permissions = json.dumps(collaborator.permissions.raw_data, indent = 4)
                    json_permissions = json_permissions.replace('{', '')
                    json_permissions = json_permissions.replace('}', '')
                    worksheet.write(row, rows["collaborators"], collaborator.login)
                    worksheet.write(row, rows["repo_admin"], repo.get_collaborator_permission(collaborator))
                    worksheet.write(row, rows["permissions"], json_permissions)
                    row += 1


def search_repo_files(g, o, xlsx):
    rate_limit = g.get_rate_limit()
    rate = rate_limit.search
    if rate.remaining == 0:
        print(f'You have 0/{rate.limit} API calls remaining. Reset time: {rate.reset}')
        return
    else:
        print(f'You have {rate.remaining}/{rate.limit} API calls remaining')

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for repo in org.get_repos():
                print(f"{repo.name}")
                filename = "requirements.yml"
                keyword = "ansible-role-gts-cm-upload-results.git"
                q = "NOT ansible_role_sfs_upload.git filename:requirements.yml org:uuc"
                query = f"{keyword} in:filename:{filename} path:/ repo:{repo.name}"
                results = g.search_code(q, order='desc')
                for result in results:
                    print(f"{result}")
                print(f"\n\n\n")

def main():
    # Setup.
    load_dotenv()
    GIT_TOKEN = os.environ.get("GIT_TOKEN", None)
    if GIT_TOKEN is None:
        print("No access token provided")
        exit(1)
    hostname = "github.kyndryl.net"
    base_url = f"https://{hostname}/api/v3"
    g = Github(base_url=base_url, login_or_token=GIT_TOKEN)

    xlsx = get_xlsx_obj()
    #orgs = get_orgs_list(g)
    # Instead of iterating orgs, we know that we need only "uuc".
    # The get_orgs_repo_details() has been modified accordingly. 
    #get_orgs_repo_details(g, "uuc", xlsx)

    # Get teams permissions for each repo.
    #get_teams_repo_permissions(g, "uuc", xlsx)

    # Get members permissions for each repo.
    #get_memebers_repo_permissions(g, "uuc", xlsx)

    # Get requirements.txt from each repo.
    search_repo_files(g, "uuc", xlsx)

    save_and_close_xlsx_obj(xlsx)


if __name__ == '__main__':
    start_time = time.time()
    main()
    print(f"Total execution time: {time.time() - start_time} seconds")
