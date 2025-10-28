const getMyFilters = async () => {
    const userId = Number(document.getElementById('userId').value);
    const res = await fetch('http://localhost:4000/graphql', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            query: `
                query {
                    filtersByUser(user_id: ${userId}) {
                        id,
                        name,
                        category,
                        dateFrom,
                        dateTo,
                        deleted,
                        user_id
                    }
                }
            `,
        }),
    });
    const data = await res.json();
    return data.data.filtersByUser;
}
const populateSavedFilters = (filters) => {
    const select = document.getElementById('saved-filters');
    filters.forEach(filter => {
        const option = document.createElement('option');
        option.value = filter.id;
        option.textContent = filter.name;
        select.appendChild(option);
    });
}

getMyFilters().then(populateSavedFilters);

const saveFilter = async () => {
    const idFilter = Number(document.getElementById('saved-filters').value);

    const userId = Number(document.getElementById('userId').value);

    let category = document.getElementById('category').value;
    if (category === 'TODAS' || category === '') category = null;

    const dateFrom = document.getElementById('dateFrom').value || null;
    const dateTo = document.getElementById('dateTo').value || null;

    let deleted = null;
    const selectedDeleted = document.getElementById('deleted').value;
    if(selectedDeleted === '') deleted = null;
    else if(selectedDeleted === 'false') deleted = false;
    else if(selectedDeleted === 'true') deleted = true;

    const mutationInput = { user_id: userId };
    if (category !== null) mutationInput.category = category;
    if (dateFrom !== null) mutationInput.dateFrom = dateFrom;
    if (dateTo !== null) mutationInput.dateTo = dateTo;
    if (deleted !== null) mutationInput.deleted = deleted;
    
    //si el id esta presente usamos la mutacion updateFilter d: caso contrario gusrdamos el nuevo filtro
    if (idFilter > 0) {
        //aca ejecuto esta query para conseguir el nombre original y no modificarlo
        try {
            const res = await fetch('http://localhost:4000/graphql', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    query: `
                        query ($id: ID!) {
                            filterById(id: $id) {
                                id,
                                name
                            }
                        }
                    `,
                    variables: { id: idFilter }
                }),
            });
            
            const data = await res.json();
            const originalName = data.data.filterById.name;
            mutationInput.name = originalName;
        } catch(error) {
            showError("Error al obtener el nombre del filtro original.");
            return;
        }
        try {
            await fetch('http://localhost:4000/graphql', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    query: `
                        mutation ($id: ID!, $input: DonationFilterInput!) {
                            updateFilter(id: $id, input: $input) {
                                id,
                                name
                            }
                        }
                    `,
                    variables: { id: idFilter, input: mutationInput }
                }),
            });
        } catch(error) {
            showError("Error al actualizar el filtro. Intente nuevamente.");
            return;
        }
        showSuccess("Filtro actualizado exitosamente.");
        return;
    }

    let filterName = prompt("Ingrese un nombre para el filtro:");
    if (!filterName) {
        showError("El nombre del filtro no puede estar vacío.");
        return;
    }
    const name = filterName.trim();
    mutationInput.name = name;
    try {
        await fetch('http://localhost:4000/graphql', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                query: `
                    mutation ($input: DonationFilterInput!) {
                        addFilter(input: $input) {
                            id,
                            name
                        }
                    }
                `,
                variables: { input: mutationInput }
            }),
        });
    } catch (error) {
        showError("Error al guardar el filtro. Intente nuevamente.");
    }
    showSuccess("Filtro guardado exitosamente.");
};

document.getElementById('save-filters').addEventListener('click', saveFilter);

document.getElementById('saved-filters').addEventListener('change', async (event) => {
    const filterId = event.target.value;
    if (!filterId) return;
    const res = await fetch('http://localhost:4000/graphql', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            query: `
                query ($id: ID!) {
                    filterById(id: $id) {
                        id,
                        name,
                        category,
                        dateFrom,
                        dateTo,
                        deleted,
                        user_id
                    }
                }
            `,
            variables: { id: filterId }
        }),
    });
    const data = await res.json();
    const filter = data.data.filterById;
    document.getElementById('category').value = filter.category || 'TODAS';
    if (filter.dateFrom) {
        const [day, month, year] = filter.dateFrom.split('/');
        const parsedDate = new Date(`${year}-${month}-${day}`);
        filter.dateFrom = parsedDate.toISOString().split('T')[0];
    } else {
        filter.dateFrom = null;
    }
    if (filter.dateTo) {
        const [day, month, year] = filter.dateTo.split('/');
        const parsedDate = new Date(`${year}-${month}-${day}`);
        filter.dateTo = parsedDate.toISOString().split('T')[0];
    } else {
        filter.dateTo = null;
    }
    document.getElementById('dateFrom').value = filter.dateFrom;
    document.getElementById('dateTo').value = filter.dateTo;
    document.getElementById('deleted').value = filter.deleted === null ? '' : filter.deleted.toString();
});

const deleteErrorDiv = (errorDiv) => {
    setTimeout(() => errorDiv.remove(), 5000);
}

const showError = (message) => {
    if (document.getElementById('errorDiv') != null) {
        document.getElementById('errorDiv').remove();
    }
    let errorDiv = document.createElement('div');
    errorDiv.id = "errorDiv";
    errorDiv.className = "error";
    errorDiv.innerHTML = message;
    document.getElementById('error').appendChild(errorDiv);
    deleteErrorDiv(errorDiv);
}

const showSuccess = (message) => {
    if (document.getElementById('sucessDiv') != null) {
        document.getElementById('sucessDiv').remove();
    }
    let sucessDiv = document.createElement('div');
    sucessDiv.id = "sucessDiv";
    sucessDiv.className = "success";
    sucessDiv.innerHTML = message;
    document.getElementById('success').appendChild(sucessDiv);
    deleteErrorDiv(sucessDiv);
}

 document.getElementById("generate-excel").addEventListener("click", async function(event) {
    event.preventDefault();

    try {
        // Determinar el tipo de exportación desde el atributo data-export-type del botón
        const btn = document.getElementById("generate-excel");
        const exportType = btn ? btn.dataset.exportType : 'all';

        // Construir URL con query param si corresponde
        let fetchUrl = "http://localhost:9093/api/export/donations";
        if (exportType && exportType !== 'all') {
            fetchUrl += `?type=${encodeURIComponent(exportType)}`;
        }

        const response = await fetch(fetchUrl, {
            method: "GET"
        });

        if (!response.ok) {
            throw new Error(`Error al generar el Excel: ${response.status} ${response.statusText}`);
        }

        const blob = await response.blob();

        // Intentar obtener filename real desde el header Content-Disposition
        let filename;
        try {
            const contentDisposition = response.headers.get('Content-Disposition') || response.headers.get('content-disposition');
            if (contentDisposition) {
                const match = /filename\*=UTF-8''([^;\n]+)|filename="?([^";\n]+)"?/.exec(contentDisposition);
                if (match) {
                    filename = match[1] ? decodeURIComponent(match[1]) : match[2];
                }
            }
        } catch (err) {
            console.warn('No se pudo parsear Content-Disposition:', err);
        }

        if (!filename) {
            // Fallback según tipo
            if (exportType === 'received') filename = 'donaciones_recibidas.xlsx';
            else if (exportType === 'made') filename = 'donaciones_realizadas.xlsx';
            else filename = 'donaciones.xlsx';
        }

        const objectUrl = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = objectUrl;
        a.download = filename;

        document.body.appendChild(a);
        a.click();

        a.remove();
        setTimeout(() => window.URL.revokeObjectURL(objectUrl), 1000);

    } catch (error) {
        console.error("Error descargando el Excel:", error);
        alert("No se pudo descargar el archivo. Revisa la consola para más detalles.");
    }
});
