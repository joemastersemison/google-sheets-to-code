import type { DependencyNode, ParsedFormula, Sheet } from "../types/index.js";

export class DependencyAnalyzer {
  dependencies: Map<string, DependencyNode> = new Map();

  buildDependencyGraph(
    sheets: Map<string, Sheet>
  ): Map<string, DependencyNode> {
    this.dependencies.clear();

    // First pass: Create nodes for all cells with formulas
    for (const [sheetName, sheet] of sheets) {
      for (const [cellRef, cell] of sheet.cells) {
        if (cell.formula && cell.parsedFormula) {
          const nodeId = `${sheetName}!${cellRef}`;
          this.dependencies.set(nodeId, {
            cellRef,
            sheetName,
            dependencies: new Set(),
            formula: cell.parsedFormula,
          });
        }
      }
    }

    // Second pass: Extract dependencies from formulas
    for (const [_nodeId, node] of this.dependencies) {
      if (node.formula) {
        const refs = this.extractReferences(node.formula, node.sheetName);
        node.dependencies = refs;
      }
    }

    // Check for circular dependencies
    this.detectCircularDependencies();

    return this.dependencies;
  }

  private extractReferences(
    formula: ParsedFormula,
    currentSheet: string
  ): Set<string> {
    const references = new Set<string>();

    const traverse = (node: ParsedFormula) => {
      if (node.type === "reference") {
        const ref = this.normalizeReference(node.value, currentSheet);
        references.add(ref);
      }

      if (node.children) {
        for (const child of node.children) {
          traverse(child);
        }
      }
    };

    traverse(formula);
    return references;
  }

  private normalizeReference(ref: string, currentSheet: string): string {
    // Handle sheet references
    if (ref.includes("!")) {
      return ref;
    }

    // Handle range references
    if (ref.includes(":")) {
      const [start, end] = ref.split(":");
      const normalizedStart = this.normalizeCellRef(start);
      const normalizedEnd = this.normalizeCellRef(end);
      return `${currentSheet}!${normalizedStart}:${normalizedEnd}`;
    }

    // Single cell reference
    return `${currentSheet}!${this.normalizeCellRef(ref)}`;
  }

  private normalizeCellRef(ref: string): string {
    // Remove absolute reference markers ($)
    return ref.replace(/\$/g, "");
  }

  getCalculationOrder(): string[] {
    const visited = new Set<string>();
    const order: string[] = [];
    const visiting = new Set<string>();

    const visit = (nodeId: string) => {
      if (visited.has(nodeId)) return;
      if (visiting.has(nodeId)) {
        throw new Error(`Circular dependency detected involving ${nodeId}`);
      }

      visiting.add(nodeId);
      const node = this.dependencies.get(nodeId);

      if (node) {
        for (const dep of node.dependencies) {
          // Only visit dependencies that are cells with formulas
          if (this.dependencies.has(dep)) {
            visit(dep);
          }
        }
      }

      visiting.delete(nodeId);
      visited.add(nodeId);
      order.push(nodeId);
    };

    // Visit all nodes
    for (const nodeId of this.dependencies.keys()) {
      visit(nodeId);
    }

    return order;
  }

  private detectCircularDependencies() {
    const visited = new Set<string>();
    const recursionStack = new Set<string>();
    const path: string[] = [];

    const detectCycle = (nodeId: string): boolean => {
      visited.add(nodeId);
      recursionStack.add(nodeId);
      path.push(nodeId);

      const node = this.dependencies.get(nodeId);
      if (node) {
        for (const dep of node.dependencies) {
          if (!visited.has(dep)) {
            if (detectCycle(dep)) {
              return true;
            }
          } else if (recursionStack.has(dep)) {
            // Found a cycle
            const cycleStart = path.indexOf(dep);
            const cycle = path.slice(cycleStart).concat(dep);
            throw new Error(
              `Circular dependency detected: ${cycle.join(" -> ")}`
            );
          }
        }
      }

      path.pop();
      recursionStack.delete(nodeId);
      return false;
    };

    for (const nodeId of this.dependencies.keys()) {
      if (!visited.has(nodeId)) {
        detectCycle(nodeId);
      }
    }
  }

  getDependents(cellRef: string): Set<string> {
    const dependents = new Set<string>();

    for (const [nodeId, node] of this.dependencies) {
      if (node.dependencies.has(cellRef)) {
        dependents.add(nodeId);
      }
    }

    return dependents;
  }

  getTransitiveDependencies(cellRef: string): Set<string> {
    const result = new Set<string>();
    const visited = new Set<string>();

    const collect = (ref: string) => {
      if (visited.has(ref)) return;
      visited.add(ref);

      const node = this.dependencies.get(ref);
      if (node) {
        for (const dep of node.dependencies) {
          result.add(dep);
          collect(dep);
        }
      }
    };

    collect(cellRef);
    return result;
  }

  getTransitiveDependents(cellRef: string): Set<string> {
    const result = new Set<string>();
    const visited = new Set<string>();

    const collect = (ref: string) => {
      if (visited.has(ref)) return;
      visited.add(ref);

      const dependents = this.getDependents(ref);
      for (const dep of dependents) {
        result.add(dep);
        collect(dep);
      }
    };

    collect(cellRef);
    return result;
  }
}
