type: patch

**Qué.** Después del merge, `safe_merge.py` busca el worktree enlazado que tiene abierta la rama local y se lo pasa a `safe_cleanup`. El worktree y la rama se retiran solo si el worktree está limpio, no tiene procesos dentro y la punta está preservada. Si `safe_merge` corre dentro de ese worktree, o la retirada no es segura, lo dice e imprime una línea `PENDING:` con los comandos exactos.
**Causa.** `git branch -D` fallaba con `used by worktree`, y el coordinador tenía que terminar la limpieza a mano (access2web-blueprint #1577, #1501).
**Arreglo.** `branch_worktree` lee `git worktree list --porcelain` y `cleanup_local_branch` llama a `safe_cleanup.cleanup(worktree=...)`. La rama remota se sigue conservando siempre.
