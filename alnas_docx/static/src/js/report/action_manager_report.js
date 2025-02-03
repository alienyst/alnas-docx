odoo.define('alnas_docx.report', function (require) {
    "use strict";

    var core = require("web.core");
    var ActionManager = require("web.ActionManager");
    var framework = require("web.framework");
    var session = require("web.session");
    var _t = core._t;

    ActionManager.include({
        _downloadReportDocxtpl: function (url, actions) {
            var self = this;
            framework.blockUI();
            var type = "docxtpl";
            var cloned_action = _.clone(actions);
            var new_url = url;

            if (
                _.isUndefined(cloned_action.data) ||
                _.isNull(cloned_action.data) ||
                (_.isObject(cloned_action.data) && _.isEmpty(cloned_action.data))
            ) {
                if (cloned_action.context.active_ids) {
                    new_url += "/" + cloned_action.context.active_ids.join(",");
                }
            } else {
                new_url +=
                    "?options=" +
                    encodeURIComponent(JSON.stringify(cloned_action.data));
                new_url +=
                    "&context=" +
                    encodeURIComponent(JSON.stringify(cloned_action.context));
            }

            console.log(new_url)

            return new Promise(function (resolve, reject) {
                var blocked = !session.get_file({
                    url: new_url,
                    data: {
                        data: JSON.stringify([new_url, type]),
                        context: JSON.stringify(cloned_action.context),
                    },
                    success: resolve,
                    error: (error) => {
                        self.call("crash_manager", "rpc_error", error);
                        reject();
                    },
                    complete: framework.unblockUI,
                });
                if (blocked) {
                    var message = _t(
                        "A popup window with your report was blocked. You " +
                            "may need to change your browser settings to allow " +
                            "popup windows for this page."
                    );
                    self.do_warn(_t("Warning"), message, true);
                }
            });
        },

        _triggerDownload: function (action, options, type) {
            var self = this;
            var reportUrls = this._makeReportUrls(action);
            if (type === "docxtpl") {
                return this._downloadReportDocxtpl(reportUrls[type], action).then(
                    function () {
                        if (action.close_on_report_download) {
                            var closeAction = {type: "ir.actions.act_window_close"};
                            return self.doAction(
                                closeAction,
                                _.pick(options, "on_close")
                            );
                        }
                        return options.on_close();
                    }
                );
            }
            return this._super.apply(this, arguments);
        },

        _makeReportUrls: function (action) {
            var reportUrls = this._super.apply(this, arguments);
            reportUrls.docxtpl = "/report/docxtpl/" + action.report_name;
            return reportUrls;
        },

        _executeReportAction: function (action, options) {
            var self = this;
            if (action.report_type === "docxtpl") {
                return self._triggerDownload(action, options, "docxtpl");
            }
            return this._super.apply(this, arguments);
        },
    });
});