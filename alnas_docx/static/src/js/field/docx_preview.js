/** @odoo-module **/

import {
    Component,
    onMounted,
    onPatched,
    onWillStart,
    onWillUnmount,
    useRef,
    useState,
} from "@odoo/owl";
import {registry} from "@web/core/registry";
import {_t} from "@web/core/l10n/translation";
import {loadBundle} from "@web/core/assets";
import {
    ReferenceField,
    referenceField,
} from "@web/views/fields/reference/reference_field";

export class DocxPreviewField extends Component {
    static template = "alnas_docx.DocxPreviewField";
    static components = {ReferenceField};
    static props = ReferenceField.props;
    static defaultProps = ReferenceField.defaultProps;

    setup() {
        this.preview = useRef("preview");
        this.previewState = useState({loading: false, error: null});
        this.previewKey = null;
        this.previewVersion = 0;
        onWillStart(() => loadBundle("alnas_docx.docx_preview"));
        onMounted(() => this.updatePreview());
        onPatched(() => this.updatePreview());
        onWillUnmount(() => ++this.previewVersion);
    }

    async updatePreview() {
        const value = this.props.record.data[this.props.name];
        const configId = this.props.record.resId;
        const key = value
            ? `${configId || "new"}:${value.resModel}:${value.resId}`
            : "";
        if (key === this.previewKey) {
            return;
        }

        this.previewKey = key;
        const version = ++this.previewVersion;
        this.previewState.error = null;
        this.previewState.loading = Boolean(value);

        if (!value) {
            return;
        }
        if (!configId) {
            this.previewState.loading = false;
            this.previewState.error = _t("Save the report configuration before previewing.");
            return;
        }

        try {
            const response = await fetch(
                `/docx-preview/${configId}/${value.resId}`
            );
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(errorText || `Preview request failed (${response.status})`);
            }
            if (!window.docx) {
                throw new Error(_t("DOCX preview library is unavailable"));
            }

            const rendered = document.createElement("div");
            await window.docx.renderAsync(await response.arrayBuffer(), rendered);
            if (version === this.previewVersion) {
                this.preview.el.replaceChildren(...rendered.childNodes);
            }
        } catch (error) {
            if (version === this.previewVersion) {
                this.previewState.error = error.message;
            }
        } finally {
            if (version === this.previewVersion) {
                this.previewState.loading = false;
            }
        }
    }
}

registry.category("fields").add("docx_preview", {
    ...referenceField,
    component: DocxPreviewField,
});
