import {address, coords, checkBounds, attr} from './xml.js';

// Keep literals, names, sheet qualifiers and structured references opaque.
export function translateShared(formula, master, target) {
  const [mr, mc] = coords(master), [tr, tc] = coords(target);
  const tokens = /"(?:[^"]|"")*"|'(?:[^']|'')*'|\[|[\p{L}\p{N}_.$\\]+|./gsuy;
  let result='', match;
  while ((match=tokens.exec(formula))) {
    if (match[0]==='[') {
      // A structured reference can contain nested brackets and escaped punctuation.
      let depth=1, i=tokens.lastIndex;
      while (i<formula.length && depth) {
        const ch=formula[i++];
        if (ch==="'" && /['[\]#@]/.test(formula[i]||' ')) i++;
        else if (ch==='[') depth++;
        else if (ch===']') depth--;
      }
      result+=formula.slice(match.index,i);tokens.lastIndex=i;
    }else result+=shift(match[0],match.index);
  }
  return result;
  function shift(word,offset) {
    const before = formula.slice(0, offset).trimEnd(), after = formula.slice(offset + word.length).trimStart();
    if (after.startsWith('!') || after.startsWith('(') || after.startsWith('[')) return word;
    if (!before.endsWith('!') && /^:\s*(?:'(?:[^']|'')*'|[\w.]+)!/.test(after)) return word;
    const cell = /^(\$?)([a-z]{1,3})(\$?)([1-9]\d*)$/i.exec(word);
    const axis = !cell && (before.endsWith(':') || after.startsWith(':')) && /^(\$?)([a-z]{1,3}|[1-9]\d*)$/i.exec(word);
    if (!cell && !axis) return word;
    const letters = cell?.[2] || (/[a-z]/i.test(axis[2]) ? axis[2] : '');
    const digits = cell?.[4] || (!letters ? axis[2] : '');
    let col = letters ? coords(letters.toUpperCase()+'1')[1] : 0, row = Number(digits);
    if (col > 16384 || row > 1048576) return word; // May be a defined name.
    const colLock = (cell || axis)[1], rowLock = cell ? cell[3] : axis[1];
    if (letters && !colLock) col += tc-mc;
    if (digits && !rowLock) row += tr-mr;
    if ((letters && (col<1 || col>16384)) || (digits && (row<1 || row>1048576))) return '#REF!';
    return (letters ? colLock+address(1,col).slice(0,-1) : '') + (digits ? rowLock+row : '');
  }
}

// A local worksheet qualifier can be quoted even when the name is a simple word.
export function renameQualifier(expression,oldName,newName) {
  const token=/"(?:[^"]|"")*"|'(?:[^']|'')*'!|(?:\[[^\]]*\])?[\p{L}\p{N}_.]+!/gu;
  return expression.replace(token,word=>{
    if(!word.endsWith('!'))return word;
    const name=word.slice(0,-1).replace(/^'|'$/g,'').replaceAll("''","'");
    return name.toUpperCase()===oldName.toUpperCase()?"'"+newName.replaceAll("'","''")+"'!":word;
  });
}

export function formulaGroup(f, anchor) {
  const ref = (attr(f,'ref') || anchor).replaceAll('$','').toUpperCase();
  const [start, end=start] = ref.split(':');
  const [top,left] = coords(start), [bottom,right] = coords(end);
  checkBounds(top,left); checkBounds(bottom,right);
  if (bottom<top || right<left) throw Error('Invalid formula range '+ref+'.');
  const yes = name => ['1','true'].includes(attr(f,name));
  const type = attr(f,'t'), two = yes('dt2D'), row = yes('dtr');
  const group = {type, anchor, ref, top, left, bottom, right, expression:f.textContent,
    label:type==='array'?'array':type==='shared'?'shared formula':two?'two-input what-if table':'one-input what-if table'};
  if (type==='dataTable') Object.assign(group, {two, row, input1:attr(f,'r1').replaceAll('$','').toUpperCase(),
    input2:attr(f,'r2').replaceAll('$','').toUpperCase(), deleted:yes('del1')||yes('del2')});
  // A rectangular reference returns its elements directly when the shapes agree.
  if (type==='array') {
    const match = /^((?:'(?:[^']|'')*'|[\w.]+)!)?(\$?[A-Z]{1,3}\$?[1-9]\d*):(\$?[A-Z]{1,3}\$?[1-9]\d*)$/i.exec(f.textContent);
    if (match) {
      const at = s=>coords(s.replaceAll('$','').toUpperCase());
      const [r,c]=at(match[2]), [lastR,lastC]=at(match[3]);
      if (lastR-r===bottom-top && lastC-c===right-left) group.source={r,c, qualifier:match[1]||''};
    }
  }
  return group;
}

export function installGroups(ctx, sheet, groups) {
  for (const group of groups) {
    for (let r=group.top;r<=group.bottom;r++) for (let c=group.left;c<=group.right;c++) {
      const key=address(r,c);
      if (!sheet.cells[key]) ctx.put(sheet,key,{v:''});
      const cell=sheet.cells[key];
      if (cell.group) throw Error('Overlapping formula ranges at '+sheet.name+'!'+key+'.');
      cell.group=group;
    }
  }
  sheet.groups=groups;
}
