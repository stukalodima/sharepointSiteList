from azure.identity import ClientSecretCredential
from django.shortcuts import render
from msgraph import GraphServiceClient
from msgraph.generated.sites.item.lists.item.items.items_request_builder import ItemsRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration

from office365.graph_client import GraphClient

scopes = ['https://graph.microsoft.com/.default']

# Values from app registration
tenant_id = '758e0274-4556-419d-be32-fddec46d6a04'
client_id = '30f9999e-20ce-4e18-bb00-cb03b7899ee3'
client_secret = 'KwL8Q~gQ~WDozL3w.kn4a7ovi2uRDFaAxMl~Vcqb'
url = 'https://smartholdingcom.sharepoint.com/'


def index(request):
    return render(request, 'main/index.html')


async def sites(request):
    data = {
        'title': 'Список сайтів',
        'sites': []
    }

    # client = GraphClient(acquire_token_func)
    # site_list = client.sites.get().execute_query()
    #
    # data['sites'] = site_list

    graph_client = get_client()

    site_list = await graph_client.sites.get_all_sites.get()

    for site in site_list.value:
        if not site.is_personal_site:

            query_params = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters(
                expand=["fields"],
            )

            request_configuration = RequestConfiguration(
                query_parameters=query_params,
            )

            res = await graph_client.sites.by_site_id(site.id).lists.by_list_id('User Information List').items.get(
                request_configuration=request_configuration)

            users = []

            for r in res.value:
                if r.fields.additional_data.get('ContentType') == 'Person' and r.fields.additional_data.get(
                        'Title') != 'NT Service\spsearch' and r.fields.additional_data.get(
                    'Title') != 'SharePoint App' and r.fields.additional_data.get(
                    'Title') != 'System Account' and r.fields.additional_data.get(
                    'Title') != 'NT Service\SPSearch' and r.fields.additional_data.get(
                    'Title') != 'Site Collection Adm':
                    users.append(r.fields.additional_data.get('Title'))

            site_with_permissions = {
                'site': site,
                'permissions': users
            }

            data['sites'].append(site_with_permissions)

    return render(request, 'main/sites.html', data)


import msal


def acquire_token_func():
    """
    Acquire token via MSAL
    """
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id=f'{client_id}',
        client_credential=f'{client_secret}'
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token


def get_client():
    # azure.identity.aio
    credential = ClientSecretCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret)

    graph_client = GraphServiceClient(credential, scopes)

    return graph_client  # type: ignore
