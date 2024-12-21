let selectedRow = null;

// Salvar contatos no localStorage
function saveContacts() {
    const table = document.getElementById("contacts-table").getElementsByTagName('tbody')[0];
    const rows = Array.from(table.rows).map(row => ({
        name: row.cells[0].innerText,
        email: row.cells[1].innerText,
        phone: row.cells[2].innerText,
        vehicle: row.cells[3].innerText,
        observations: row.cells[4].innerText,
    }));
    localStorage.setItem("contacts", JSON.stringify(rows));
}

// Carregar contatos do localStorage
function loadContacts() {
    const table = document.getElementById("contacts-table").getElementsByTagName('tbody')[0];
    const storedContacts = JSON.parse(localStorage.getItem("contacts")) || [];
    storedContacts.forEach(contact => {
        const newRow = table.insertRow();
        newRow.onclick = () => selectRow(newRow); // Adicionar evento de seleção
        newRow.insertCell(0).innerText = contact.name;
        newRow.insertCell(1).innerText = contact.email;
        newRow.insertCell(2).innerText = contact.phone;
        newRow.insertCell(3).innerText = contact.vehicle;
        newRow.insertCell(4).innerText = contact.observations;
    });
}

// Adicionar um novo contato
function addContact() {
    const name = document.getElementById("name").value;
    const email = document.getElementById("email").value;
    const phone = document.getElementById("phone").value;
    const vehicle = document.getElementById("vehicle").value;
    const observations = document.getElementById("observations").value;

    if (name && email && phone) {
        const table = document.getElementById("contacts-table").getElementsByTagName('tbody')[0];
        const newRow = table.insertRow();
        newRow.onclick = () => selectRow(newRow); // Adicionar evento de seleção

        newRow.insertCell(0).innerText = name;
        newRow.insertCell(1).innerText = email;
        newRow.insertCell(2).innerText = phone;
        newRow.insertCell(3).innerText = vehicle;
        newRow.insertCell(4).innerText = observations;

        document.getElementById("contact-form").reset();
        saveContacts();
    } else {
        alert("Por favor, preencha todos os campos.");
    }
}

// Selecionar uma linha
function selectRow(row) {
    if (selectedRow) selectedRow.style.backgroundColor = ""; // Limpar seleção anterior
    selectedRow = row;
    selectedRow.style.backgroundColor = "#d3d3d3"; // Estilo de seleção
}

// Exportar contato selecionado para PDF
function exportSelectedToPDF() {
    if (!selectedRow) {
        alert("Por favor, selecione um contato.");
        return;
    }

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    const cells = Array.from(selectedRow.cells).map(cell => cell.innerText);
    pdf.text("Contato Selecionado", 10, 10);

    pdf.autoTable({
        head: [["Nome", "Email", "Telefone", "Veículo", "Observações"]],
        body: [cells],
        startY: 20,
    });

    pdf.save("Contato_Selecionado.pdf");
}

// Exportar todos os contatos para PDF
function exportPDF() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    pdf.text("Lista de Contatos", 10, 10);
    pdf.autoTable({
        html: "#contacts-table",
        startY: 20,
        headStyles: { fillColor: [192, 0, 0] },
    });

    pdf.save("Contatos_Nissan.pdf");
}

// Exportar todos os contatos para Excel
function exportExcel() {
    const table = document.getElementById("contacts-table");
    const data = Array.from(table.rows).map(row => Array.from(row.cells).map(cell => cell.innerText));
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Contatos");
    XLSX.writeFile(workbook, "Contatos_Nissan.xlsx");
}

// Exportar todos os contatos para Word
function exportToWord() {
    const table = document.getElementById("contacts-table");
    const rows = Array.from(table.rows);
    let content = `<h1>Lista de Contatos</h1><table border="1" style="border-collapse:collapse;">`;

    rows.forEach(row => {
        content += "<tr>";
        Array.from(row.cells).forEach(cell => {
            content += `<td style="padding:8px;">${cell.innerText}</td>`;
        });
        content += "</tr>";
    });

    content += "</table>";

    const blob = new Blob(['\ufeff' + content], {
        type: "application/msword"
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Contatos_Nissan.doc";
    link.click();
}

// Imprimir a página
function printPage() {
    const tableHTML = document.getElementById("contacts-table").outerHTML;
    const newWindow = window.open("");
    newWindow.document.write("<h1>Lista de Contatos</h1>");
    newWindow.document.write(tableHTML);
    newWindow.document.close();
    newWindow.print();
}

// Criar reunião
function openSchedulePopup() {
    const existingPopup = document.getElementById("schedule-popup");
    if (existingPopup) existingPopup.remove();

    const popup = document.createElement("div");
    popup.id = "schedule-popup";
    popup.innerHTML = `
        <div class="popup-content">
            <h2>Criar Reunião</h2>
            <input type="text" id="meeting-title" placeholder="Título da Reunião" required>
            <label for="meeting-start">Início:</label>
            <input type="datetime-local" id="meeting-start" required>
            <label for="meeting-end">Fim:</label>
            <input type="datetime-local" id="meeting-end" required>
            <textarea id="meeting-description" placeholder="Descrição (Opcional)"></textarea>
            <button onclick="saveMeeting()">Salvar</button>
            <button onclick="closePopup()">Cancelar</button>
        </div>
    `;
    document.body.appendChild(popup);
}

// Salvar reunião
function saveMeeting() {
    const title = document.getElementById("meeting-title").value;
    const start = document.getElementById("meeting-start").value;
    const end = document.getElementById("meeting-end").value;
    const description = document.getElementById("meeting-description").value;

    if (!title || !start || !end) {
        alert("Por favor, preencha todos os campos obrigatórios.");
        return;
    }

    const startDate = new Date(start).toISOString().replace(/[-:]/g, "").split(".")[0];
    const endDate = new Date(end).toISOString().replace(/[-:]/g, "").split(".")[0];

    const icsContent = `
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Nissan Calendar//PT
BEGIN:VEVENT
UID:${Date.now()}@nissan.com
DTSTAMP:${new Date().toISOString().replace(/[-:]/g, "").split(".")[0]}
DTSTART:${startDate}
DTEND:${endDate}
SUMMARY:${title}
DESCRIPTION:${description}
END:VEVENT
END:VCALENDAR
    `.trim();

    const blob = new Blob([icsContent], { type: "text/calendar" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title.replace(/\s+/g, "_")}_meeting.ics`;
    link.click();

    alert("Arquivo de reunião criado e baixado com sucesso!");
    closePopup();
}

// Fechar pop-up
function closePopup() {
    const popup = document.getElementById("schedule-popup");
    if (popup) popup.remove();
}

// Limpar todos os contatos
function clearContacts() {
    if (confirm("Tem certeza que deseja limpar todos os contatos?")) {
        document.querySelector("#contacts-table tbody").innerHTML = "";
        localStorage.removeItem("contacts");
    }
}

// Carregar contatos ao abrir a página
window.onload = loadContacts;