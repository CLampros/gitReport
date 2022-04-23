from textwrap import wrap
from webbrowser import get
from github import Github
import xlsxwriter
import time
import json

def get_xlsx_obj():
    # Create a workbook and add a worksheet.
    timestr = time.strftime("%Y%m%d-%H%M%S")
    workbook = xlsxwriter.Workbook(f'gitReport-{timestr}.xlsx')

    return workbook

def save_and_close_xlsx_obj(workbook):
    # Save and close xlsx
    workbook.close()

def get_orgs_list(g):
    orgslist = [orgs.login for orgs in g.get_user().get_orgs()]
        
    return orgslist

def get_orgs_repo_details(g, o, xlsx):
    # create a worksheet
    worksheet = xlsx.add_worksheet('UUC REPO BRANCHES')
    # formats
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

    worksheet.set_column('B:B', 50, fcenter)
    worksheet.set_column('C:C', 10, fwrap)
    worksheet.set_column('D:D', 30, fwrap)
    worksheet.set_column('E:E', 10, fwrap)
    worksheet.set_column('F:F', 10, fwrap)
    worksheet.set_column('G:G', 30, fwrap)
    worksheet.set_column('H:H', 50, fwrap)
    
    # headers
    worksheet.write(1, 1, "Repo", fheaders)
    worksheet.write(1, 2, "Private", fheaders)
    worksheet.write(1, 3, "Branches", fheaders)
    worksheet.write(1, 4, "Has Prod", fheaders)
    worksheet.write(1, 5, "Protected", fheaders)
    worksheet.write(1, 6, "Author name", fheaders)
    worksheet.write(1, 7, "Author e-mail", fheaders)

    row = 2
    for org in g.get_user().get_orgs():
        if org.login == o:
            for repo in org.get_repos():
                # name
                worksheet.write(row, 1, repo.name)

                # is private?
                worksheet.write(row, 2, repo.private)

                # brances
                branches = [branch.name for branch in repo.get_branches()]
                branches_str = "\n".join(branches)
                worksheet.write(row, 3, branches_str)

                # has prod
                has_prod = 'prod' in branches
                worksheet.write(row, 4, has_prod)

                # last commit on prod branch (if exists)
                if has_prod:
                    # is protected?
                    isProtected = repo.get_branch('prod').protected
                    worksheet.write(row, 5, isProtected)
                    
                    # commiter_name & commiter_email
                    try:
                        commiter_name = repo.get_branch('prod').commit.author.name
                        commiter_email = repo.get_branch('prod').commit.author.email
                        worksheet.write(row, 6, commiter_name)
                        worksheet.write(row, 7, commiter_email)
                    except AttributeError:
                        worksheet.write(row, 6, 'No commits')
                        worksheet.write(row, 7, 'No commits')
                else:
                    worksheet.write(row, 5, 'No prod')
                    worksheet.write(row, 6, 'No prod')
                    worksheet.write(row, 7, 'No prod')
                row += 1

def get_teams_repo_permissions(g, o, xlsx):
    # create a worksheet
    worksheet = xlsx.add_worksheet('Teams Permissions')
    # formats
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
    worksheet.set_column('D:D', 50, fwrap)
    
    # headers
    worksheet.write(1, 1, "Team", fheaders)
    worksheet.write(1, 2, "Repo", fheaders)
    worksheet.write(1, 3, "Permissions", fheaders)

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for team in org.get_teams():
                for repo in team.get_repos():
                    # team
                    worksheet.write(row, 1, team.name)

                    # repo
                    worksheet.write(row, 2, repo.name)

                    # permissions
                    json_permissions = json.dumps(team.get_repo_permission(repo).raw_data, indent = 4)
                    json_permissions = json_permissions.replace('{', '')
                    json_permissions = json_permissions.replace('}', '')
                    worksheet.write(row, 3, json_permissions)

                    row += 1

def get_memebers_repo_permissions(g, o, xlsx):
    # create a worksheet
    worksheet = xlsx.add_worksheet('Members Permissions')
    # formats
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
    worksheet.set_column('D:D', 10, fcenter)
    worksheet.set_column('E:E', 50, fwrap)
    
    # headers
    worksheet.write(1, 1, "Repo", fheaders)
    worksheet.write(1, 2, "Collaborator", fheaders)
    worksheet.write(1, 3, "Repo admin", fheaders)
    worksheet.write(1, 3, "Permissions", fheaders)

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for repo in org.get_repos():
                print(repo.name)
                for collaborator in repo.get_collaborators():
                    # https://pygithub.readthedocs.io/en/latest/github_objects/Repository.htmlgit@github.com:CLampros/gitReport.git
                    # https://docs.github.com/en/rest/reference/collaborators#get-repository-permissions-for-a-user
                    print(repo.get_collaborator_permission(collaborator))

def main():
    # setup
    hostname = "github.kyndryl.net"
    base_url = f"https://{hostname}/api/v3"
    g = Github(base_url=base_url, login_or_token="ghp_XXX")

    xlsx = get_xlsx_obj()
    orgs = get_orgs_list(g)
    # instead of iterating orgs, we know that we need only "uuc"
    # get_orgs_repo_brances has been modified accordingly 
    #get_orgs_repo_details(g, "uuc", xlsx)

    # get teams permissions for each repo
    #get_teams_repo_permissions(g, "uuc", xlsx)

    # get members permissions for each repo
    get_memebers_repo_permissions(g, "uuc", xlsx)

    save_and_close_xlsx_obj(xlsx)

if __name__ == '__main__':
    main()
