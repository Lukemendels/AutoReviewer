const fs = require('fs');
const path = require('path');
const vm = require('vm');

const htmlPath = path.resolve(__dirname, '../autoreviewer.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

// Find all script blocks
const scriptRegex = /<script>([\s\S]*?)<\/script>/g;
let match;
let jsContent = '';

// We skip the first script block (which contains JSZip) and gather the rest
let index = 0;
while ((match = scriptRegex.exec(htmlContent)) !== null) {
    if (index > 0) {
        jsContent += match[1] + '\n';
    }
    index++;
}

class MockNode {
    constructor(localName, textContent = '', parentNode = null) {
        this.localName = localName;
        this.nodeType = 1;
        this.textContent = textContent;
        this.parentNode = parentNode;
        this.childNodes = [];
    }
    appendChild(child) {
        child.parentNode = this;
        this.childNodes.push(child);
        return child;
    }
    insertBefore(newChild, refChild) {
        newChild.parentNode = this;
        const idx = this.childNodes.indexOf(refChild);
        if (idx !== -1) {
            this.childNodes.splice(idx, 0, newChild);
        } else {
            this.childNodes.push(newChild);
        }
        return newChild;
    }
    cloneNode(deep) {
        const clone = new MockNode(this.localName, this.textContent, this.parentNode);
        if (deep) {
            for (let child of this.childNodes) {
                clone.appendChild(child.cloneNode(true));
            }
        }
        return clone;
    }
    setAttributeNS(ns, name, value) {
        // Mock method
    }
    getAttribute(name) {
        return null;
    }
    getElementsByTagNameNS(ns, name) {
        const results = [];
        function traverse(n) {
            if (n.localName === name || name === '*') {
                results.push(n);
            }
            for (let child of n.childNodes) {
                traverse(child);
            }
        }
        for (let child of this.childNodes) {
            traverse(child);
        }
        return results;
    }
}

const domMock = {
    documentElement: new MockNode('p'),
    createElementNS: (ns, name) => new MockNode(name.split(':').pop())
};

const mockSandbox = {
    console: console,
    process: process,
    TextEncoder: TextEncoder,
    crypto: {
        subtle: {
            digest: async (algo, data) => {
                const crypto = require('crypto');
                return crypto.createHash('sha256').update(data).digest();
            }
        }
    },
    window: {
        addEventListener: () => {}
    },
    document: {
        getElementById: (id) => ({
            classList: { remove: () => {}, toggle: () => {}, add: () => {} },
            appendChild: () => {},
            addEventListener: () => {},
            value: '',
            textContent: ''
        }),
        addEventListener: () => {},
        querySelectorAll: () => [],
        createElement: (name) => {
            const element = {
                className: '',
                parentNode: { insertBefore: () => {} }
            };
            Object.defineProperty(element, 'textContent', {
                set: (val) => {
                    console.log(val);
                    if (val.startsWith('FAIL:')) {
                        process.exit(1);
                    }
                },
                get: () => ""
            });
            return element;
        }
    },
    DOMParser: class {
        parseFromString(str, type) {
            const pNode = new MockNode('p');
            const rNode = new MockNode('r');
            let tText = "hello world";
            const tMatch = str.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
            if (tMatch) {
                tText = tMatch[1];
            }
            const tNode = new MockNode('t', tText);
            rNode.appendChild(tNode);
            pNode.appendChild(rNode);
            return {
                documentElement: pNode,
                querySelectorAll: () => []
            };
        }
    },
    Node: {
        ELEMENT_NODE: 1
    },
    URLSearchParams: URLSearchParams
};

// Run the script in the VM
try {
    const script = new vm.Script(jsContent);
    const context = vm.createContext(mockSandbox);
    script.runInContext(context);
    
    console.log("Executing unit tests from VM...");
    context.runDiagnostics();
    console.log("All unit tests run successfully!");
} catch (e) {
    console.error("Failed to run tests:", e);
    process.exit(1);
}
