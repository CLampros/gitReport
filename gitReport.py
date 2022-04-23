from textwrap import wrap
import time
import json
import time
import calendar

from github import Github
import xlsxwriter

def get_xlsx_obj():
    # Create a workbook and add a worksheet.
    timestr = time.strftime("%Y%m%d-%H%M%S")
    workbook = xlsxwriter.Workbook(f'gitReport-{timestr}.xlsx')

    return workbook


def save_and_close_xlsx_obj(workbook):
    # Save and close xlsx
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
    worksheet.set_column('D:D', 15, fcenter)
    worksheet.set_column('E:E', 25, fwrap)
    
    # headers
    worksheet.write(1, 1, "Repo", fheaders)
    worksheet.write(1, 2, "Collaborator", fheaders)
    worksheet.write(1, 3, "Repo admin", fheaders)
    worksheet.write(1, 4, "Permissions", fheaders)

    for org in g.get_user().get_orgs():
        if org.login == o:
            row = 2
            for repo in org.get_repos():
                worksheet.write(row, 1, repo.name)
                for collaborator in repo.get_collaborators():
                    # Sleep if no more requests are available until they have been replenished
                    print(f"{row=}")
                    sleep_if_core_rate_limit_reached(g)
                    # permissions
                    # https://pygithub.readthedocs.io/en/latest/github_objects/Repository.html
                    # https://docs.github.com/en/rest/reference/collaborators#get-repository-permissions-for-a-user
                    # https://docs.github.com/en/enterprise-server@3.2/organizations/managing-access-to-your-organizations-repositories/viewing-people-with-access-to-your-repository
                    json_permissions = json.dumps(collaborator.permissions.raw_data, indent = 4)
                    json_permissions = json_permissions.replace('{', '')
                    json_permissions = json_permissions.replace('}', '')
                    worksheet.write(row, 2, collaborator.login)
                    worksheet.write(row, 3, repo.get_collaborator_permission(collaborator))
                    worksheet.write(row, 4, json_permissions)
                    row += 1


def main():
    # Setup.
    hostname = "github.kyndryl.net"
    base_url = f"https://{hostname}/api/v3"
    g = Github(base_url=base_url, login_or_token="ghp_XXX")

    # core_rate_limit = g.get_rate_limit().core
    # reset_timestamp = calendar.timegm(core_rate_limit.reset.timetuple())
    # sleep_time = reset_timestamp - calendar.timegm(time.gmtime())
    # print(f"Remaining searches ==> {g.get_rate_limit().search.remaining}")
    # print(f"{core_rate_limit.remaining=}")
    # print(f"{reset_timestamp=}")
    # print(f"{sleep_time=}")

    xlsx = get_xlsx_obj()
    #orgs = get_orgs_list(g)
    # Instead of iterating orgs, we know that we need only "uuc".
    # The get_orgs_repo_details() has been modified accordingly. 
    #get_orgs_repo_details(g, "uuc", xlsx)

    # Get teams permissions for each repo.
    #get_teams_repo_permissions(g, "uuc", xlsx)

    # Get members permissions for each repo.
    get_memebers_repo_permissions(g, "uuc", xlsx)

    save_and_close_xlsx_obj(xlsx)


if __name__ == '__main__':
    start_time = time.time()
    main()
    print(f"Total execution time: {time.time() - start_time} seconds")
