import type { DependencyNode, ParsedFormula, Sheet } from "../types/index.js";

export class DependencyAnalyzer {
  dependencies: Map<string, DependencyNode> = new Map();
  circularDependencies: Set<string> = new Set();

  buildDependencyGraph(
    sheets: Map<string, Sheet>
  ): Map<string, DependencyNode> {
    this.dependencies.clear();
    this.circularDependencies.clear();

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
      // Handle quoted sheet names: 'My Sheet'!A1
      if (ref.startsWith("'")) {
        const exclamationIndex = ref.indexOf("!");
        // Extract sheet name (including quotes) and cell reference
        const sheetPart = ref.substring(0, exclamationIndex);
        const cellPart = ref.substring(exclamationIndex + 1);

        // Remove surrounding quotes and unescape doubled quotes
        const sheetName = sheetPart.slice(1, -1).replace(/''/g, "'");

        // Normalize the cell reference part
        if (cellPart.includes(":")) {
          const [start, end] = cellPart.split(":");
          return `${sheetName}!${this.normalizeCellRef(start)}:${this.normalizeCellRef(end)}`;
        }
        return `${sheetName}!${this.normalizeCellRef(cellPart)}`;
      }

      // Handle unquoted sheet names
      const [sheetName, cellRef] = ref.split("!");
      if (cellRef.includes(":")) {
        const [start, end] = cellRef.split(":");
        return `${sheetName}!${this.normalizeCellRef(start)}:${this.normalizeCellRef(end)}`;
      }
      return `${sheetName}!${this.normalizeCellRef(cellRef)}`;
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
    // Always ensure circular dependencies are detected first
    // This is idempotent, so calling it multiple times is safe
    this.detectCircularDependencies();

    const visited = new Set<string>();
    const order: string[] = [];
    const visiting = new Set<string>();

    const visit = (nodeId: string): boolean => {
      // Already processed
      if (visited.has(nodeId)) return true;

      // Skip entirely if this cell is part of a circular dependency
      // These will be handled separately with #REF! errors
      if (this.circularDependencies.has(nodeId)) {
        visited.add(nodeId);
        // Don't add to order - circular cells are handled separately
        return false;
      }

      // Currently visiting - this means we've found an undetected cycle
      if (visiting.has(nodeId)) {
        // This shouldn't happen if detectCircularDependencies() worked correctly
        console.error(
          `⚠️  Unexpected cycle detected during topological sort at ${nodeId}`
        );
        // Mark this node and don't process it
        this.circularDependencies.add(nodeId);
        return false;
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

      // Only add to order if not circular
      if (!this.circularDependencies.has(nodeId)) {
        order.push(nodeId);
      }

      return true;
    };

    // First, visit all non-circular nodes
    for (const nodeId of this.dependencies.keys()) {
      if (!this.circularDependencies.has(nodeId)) {
        visit(nodeId);
      }
    }

    // Then add circular dependency nodes at the end
    // They'll be calculated as #REF! but we need them in the order
    for (const nodeId of this.circularDependencies) {
      if (this.dependencies.has(nodeId)) {
        order.push(nodeId);
      }
    }

    return order;
  }

  detectCircularDependencies() {
    // If we've already detected cycles, don't re-run
    // unless the dependency graph has changed
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
          // Only follow dependencies that are actual nodes in the graph (cells with formulas)
          if (!this.dependencies.has(dep)) {
            continue;
          }

          if (!visited.has(dep)) {
            if (detectCycle(dep)) {
              return true;
            }
          } else if (recursionStack.has(dep)) {
            // Found a cycle - mark all cells in the cycle
            const cycleStart = path.indexOf(dep);
            const cycle = path.slice(cycleStart);

            // Store all cells involved in this circular dependency
            for (const cellId of cycle) {
              this.circularDependencies.add(cellId);
            }

            // Log warning but don't throw - let code generation handle it
            console.warn(
              `⚠️  Circular dependency detected: ${cycle.join(" -> ")} -> ${dep}`
            );

            return true; // Signal that a cycle was found
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
