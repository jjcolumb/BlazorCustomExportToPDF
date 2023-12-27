window.BlazorDownloadFile = function (filename, data, contentType) {
    const a = document.createElement('a');
    a.href = `data:${contentType};base64,${data}`;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}