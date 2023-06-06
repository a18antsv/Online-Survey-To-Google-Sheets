export const promiseHandler = promise => {
    return promise
        .then(response => [undefined, response])
        .catch(error => [error, undefined]);
}
