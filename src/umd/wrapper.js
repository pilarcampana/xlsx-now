// Hand-written UMD envelope for the bundle described by tsconfig.umd.json.
//
// `tsc --module amd --outFile` already emits every module of src/core into one
// file, in dependency order, leaving `fflate` as an unresolved named
// dependency. What is missing is the UMD detection and an AMD runtime small
// enough to read, which is what this file adds; scripts/build-umd.ts splices
// the two together. No bundler takes part and nothing is inlined: `fflate` is
// obtained from the host at load time, just like the ESM build does.
(function (root, factory) {
    if (typeof exports === 'object' && typeof module === 'object') {
        module.exports = factory(require('fflate'));
    } else if (typeof define === 'function' && define.amd) {
        define(['fflate'], factory);
    } else {
        if (!root.fflate) {
            throw new Error('xlsx-now: fflate must be loaded before this script (e.g. <script src="https://unpkg.com/fflate"></script>)');
        }
        root.xlsxNow = factory(root.fflate);
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function (fflate) {
    // Minimal AMD runtime for the modules below. They arrive in dependency
    // order, so each `define` can run its factory right away: everything it
    // names is already in the registry. `fflate` is seeded from the host, and
    // it is the only non-local name any of them asks for.
    var registry = { fflate: fflate };
    function require(name) {
        var required = registry[name];
        if (!required) {
            throw new Error('xlsx-now: unknown module "' + name + '" in the UMD bundle');
        }
        return required;
    }
    function define(name, dependencies, factory) {
        var exports = (registry[name] = {});
        factory.apply(
            null,
            dependencies.map(function (dependency) {
                if (dependency === 'exports') return exports;
                if (dependency === 'require') return require;
                return require(dependency);
            })
        );
    }

    // <<< tsc AMD output is injected here by scripts/build-umd.ts >>>

    return require('index');
});
