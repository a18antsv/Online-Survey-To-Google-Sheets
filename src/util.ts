export const handler = <T>(promise: Promise<T>): Promise<{response?: T, error?: unknown}> => {
    return promise
        .then(response => ({response}))
        .catch(error => ({error}));
}
