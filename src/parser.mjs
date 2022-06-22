function parseToken(text) {
    let tokens = []; // [{type:'',text:''}]
    let buffer = "";
    let cmdMode = false;
    let stringMode = false
    text += "\0"; // mask RangeError
    for (let i = 0; i < text.length - 1; i++) {
        let c = text[i];
        if (!cmdMode && c != "{") {
            buffer += c;
            continue;
        }
        if (stringMode && !"\"\\".includes(c)) {
            buffer += c;
            continue;
        }
        switch (c) {
            case '{':
                if (buffer.length > 0) {
                    tokens.push({ type: 'text', text: buffer })
                    buffer = "";
                }
                tokens.push({ type: '{', text: '{' })
                if ("$@#".includes(text[i + 1])) {
                    tokens.push({ type: 'type', text: text[i + 1] })
                    i++;
                }
                cmdMode = true;
                break;
            case '}':
                if (buffer.length > 0) {
                    tokens.push({ type: 'text', text: buffer })
                    buffer = "";
                }
                tokens.push({ type: '}', text: '}' })
                cmdMode = false;
                break;
            case '(':
                tokens.push({ type: 'text', text: buffer })
                buffer = "";
                tokens.push({ type: '(', text: '(' })
                break;
            case ')':
                if (buffer.length > 0) {
                    tokens.push({ type: 'text', text: buffer })
                    buffer = "";
                }
                tokens.push({ type: ')', text: ')' })
                break;
            case '.':
                if (buffer.length > 0) {
                    tokens.push({ type: 'text', text: buffer })
                    buffer = "";
                }
                tokens.push({ type: '.', text: '.' })
                break;
            case ',':
                tokens.push({ type: 'text', text: buffer })
                buffer = "";
                tokens.push({ type: ',', text: ',' })
                break;
            case '\\':
                buffer += text[i + 1];
                i++;
                break;
            case '"':
                stringMode = !stringMode
                break
            default:
                buffer += text[i];
                break;
        }
    }
    if (buffer.length > 0) {
        if (buffer[buffer.length - 1] == '\0') {
            throw 'cell = ' + text + ' cant be "{xxx\"'
        }
        tokens.push({ type: 'text', text: buffer })
    }
    tokens.push({ type: 'end', text: null })
    return tokens;
}

function parseExps(tokens) {
    let output = [];
    let exps = [];

    let newExp = () => ({ colName: "", funcs: [], type: null, raw: "{" })
    let newFunc = () => ({ name: '', args: [] })

    let exp = null;
    let func = null;
    let nextTextCall = null;
    let slot = false;

    let cmdMode = false;
    for (const token of tokens) {
        let type = token.type;
        if (cmdMode) {
            exp.raw += token.text;
        }
        switch (type) {
            case "{":
                cmdMode = true;
                exp = newExp();
                nextTextCall = (text) => exp.colName = text;
                if (!slot) {
                    //double exp
                    output.push("")
                }
                slot = false;
                break;
            case "}":
                if (func != null) {
                    exp.funcs.push(func)
                }
                exps.push(exp);
                exp = null;
                cmdMode = false;
                break;
            case "type":
                exp.type = token.text;
                break;
            case ".":
                nextTextCall = (text) => func.name = text;
                if (func != null) {
                    exp.funcs.push(func)
                }
                func = newFunc();
                break;
            case "(":
            case ",":
                nextTextCall = (text) => func.args.push(parseType(text));
                break;
            case ")":
                nextTextCall = (text) => func.name = text;
                if (func != null) {
                    exp.funcs.push(func)
                }
                func = null
                break;
            case "text":
                if (!cmdMode) {
                    output.push(token.text)
                    slot = true;
                    break;
                }
                if (nextTextCall) {
                    nextTextCall(token.text)
                    nextTextCall = null;
                } else {
                    throw "token error"
                }
                break;
        }
    }

    return {
        output,
        exps
    }
}

function parseType(str) {
    if (str[0] === '"' && str[str.length - 1] === '"') {
        return str.substring(1, str.length - 1)
    }
    if (str === 'true') return true
    if (str === 'false') return false
    if (!Number.isNaN(Number(str))) return Number(str)
    return str
}

function parseScript(text) {
    return parseExps(parseToken(text))
}

export default parseScript