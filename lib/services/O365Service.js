var O365Service = /** @class */ (function () {
    function O365Service() {
    }
    O365Service.prototype.setup = function (context) {
        this.context = context;
    };
    O365Service.prototype.getGroupBySiteName = function (siteName) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            try {
                // Prepare the output array
                var sites = new Array();
                _this.context.msGraphClientFactory
                    .getClient()
                    .then(function (client) {
                    client
                        .api("/sites?search='" + siteName + "'")
                        .get(function (error, groups, rawResponse) {
                        // Map the response to the output array
                        groups.value.map(function (item) {
                            sites.push({
                                id: item.id,
                            });
                        });
                        console.log(sites);
                        resolve(sites);
                    });
                });
            }
            catch (error) {
                console.error(error);
            }
        });
    };
    return O365Service;
}());
export { O365Service };
var GroupService = new O365Service();
export default GroupService;
//# sourceMappingURL=O365Service.js.map