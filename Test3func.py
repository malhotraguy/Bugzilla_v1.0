def bugzilla(id,include_fields_values=None):

    import requests
    import json

    id=int(id)
    url = 'https://bugzilla.mozilla.org/rest/bug'
    u = url + "/"+str(id)+"?include_fields=component,product,creation_time,last_change_time,duplicates"
    search_results = requests.get(u)
    #print(search_results.text)
    Results=json.loads(search_results.text)
    #print(Results["bugs"])
    if (include_fields_values !=None):
        return (Results["bugs"][0][include_fields_values])
    else:
        return (Results["bugs"][0])




#component,product,creation_time,last_change_time,duplicates

