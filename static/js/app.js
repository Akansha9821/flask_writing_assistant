document.addEventListener("DOMContentLoaded", () => {
    const category = document.getElementById("category");
    const details = document.querySelector('textarea[name="details"]');
    if (!category || !details) return;

    const hints = {
        leave: "Mention leave dates, reason, work handover and expected return date.",
        bank: "Mention account/service type, transaction date, reference number and requested action. Never enter your PIN or password.",
        office_followup: "Mention the earlier discussion, pending item, responsible person and expected deadline.",
        device_complaint: "Mention device model, purchase date, issue, troubleshooting attempted and warranty status.",
        exam: "Mention subject, examination date, schedule conflict, roll number and requested support.",
        project: "Mention project title, guide, deadline, current issue and required approval.",
        placement: "Mention company/role, application stage, interview date and assistance required.",
        practical: "Mention subject, practical date, reason for absence or issue and requested alternative.",
        viva: "Mention project title, scheduled date, conflict and requested rescheduling.",
        payment: "Mention amount, payment date, receipt/reference number and current status.",
        class: "Mention class/subject, date, concern and requested action.",
        customer_service: "Mention product/service, order or ticket number, issue and desired resolution."
    };

    category.addEventListener("change", () => {
        if (!details.value.trim()) details.placeholder = hints[category.value] || "Describe the issue, reason and requested action.";
    });
});
