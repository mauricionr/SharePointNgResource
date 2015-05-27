
/*

    SharePoint Module
    
    Services

*/

; (function () {
    angular.module('SharePoint', ['ngResource']);
    angular.module('SharePoint').factory('Lists', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists?:odata", null,
            {
                'update':
                {
                    method: 'POST',
                    headers:
                    {
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                    {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        'X-HTTP-Method': "",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                }
            }
            );
    }]);
    angular.module('SharePoint').factory('List', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle(':listName')?:odata", null,
            {
                'update': {
                    method: 'POST',
                    headers: {
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                    {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        'X-HTTP-Method': "",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                }
            });
    }]);
    angular.module('SharePoint').factory('ListItem', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle(':listName')/items(:itemID)?:odata", null,
            {
                'update': {
                    method: 'POST',
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                    {
                        "accept": "application/json;odata=verbose",
                        "X-HTTP-Method": "",
                        "content-type": "application/json;odata=verbose",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                }
            });
    }]);
    angular.module('SharePoint').factory('ListItems', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle(':listName')/items?:odata", null,
            {
                'update': {
                    method: 'POST',
                    headers: {
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                    {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        'X-HTTP-Method': "",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'query': {
                    isArray: false
                }
            });
    }]);
    angular.module('SharePoint').factory('ListFields', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle(':listName')/fields?:odata", null,
            {
                'update': {
                    method: 'POST',
                    headers: {
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                    {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        'X-HTTP-Method': "",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                }
            });
    }]);
    angular.module('SharePoint').factory('SiteUsers', ['$resource', function ($resource) {
        return $resource(_spPageContextInfo.webAbsoluteUrl + "/_api/web/SiteUsers?:odata", null,
            {
                'update': {
                    method: 'POST',
                    headers: {
                        "IF-MATCH": "*",
                        "content-type": "application/json;odata=verbose",
                        "X-HTTP-Method": "MERGE",
                        "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                    }
                },
                'save':
                {
                    method: 'POST',
                    headers:
                     {
                         "accept": "application/json;odata=verbose",
                         "content-type": "application/json;odata=verbose",
                         'X-HTTP-Method': "",
                         "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
                     }
                }
            });
    }]);
    angular.module('SharePoint').factory('SPUtils', function ($http, $q) {
        function sendEmail(from, to, body, subject) {
            var defer = $q.defer()
            if(to){
	            $http.defaults.headers.common.Accept = "application/json;odata=verbose";
	            $http.defaults.headers.post['Content-Type'] = 'application/json;odata=verbose';
	            $http.defaults.headers.post['X-RequestDigest'] = document.querySelector("#__REQUESTDIGEST").value;
	            var urlTemplate = _spPageContextInfo.webAbsoluteUrl + "/_api/SP.Utilities.Utility.SendEmail";
	            var emailObj = JSON.stringify({ 'properties': { '__metadata': { 'type': 'SP.Utilities.EmailProperties' }, 'From': from, 'To': { 'results': [to] }, 'Body': body, 'Subject': subject } })
	            $http.post(urlTemplate, emailObj).then(defer.resolve, defer.reject);
            }else{
            	setTimeout(function(){
            		defer.resolve()
            	},500);
            }
            return defer.promise;
        }
        return {
            sendEmail: sendEmail
        }
    });
})();

/*

	polyfills

*/
if (!String.prototype.Format) {
    String.prototype.Format = function () {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function (match, number) {
            return typeof args[number] != 'undefined'
              ? args[number]
              : match
            ;
        });
    };
}