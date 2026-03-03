'use client'

import React, { memo } from 'react'
import type { Permission } from '../hooks/usePermissions'

interface UserCardProps {
  perm: Permission
  originalIdx: number
  isAdmin: boolean
  expandedIdsIndex: number | null
  onAssignIds: (perm: Permission) => void
  onToggleExpand: (idx: number) => void
}

function UserCardInner({
  perm,
  originalIdx,
  isAdmin,
  expandedIdsIndex,
  onAssignIds,
  onToggleExpand
}: UserCardProps) {
  const level = perm.level ?? 1
  const isAdminOrSupervisor = level === 2 || level === 3

  return (
    <div className="user-card">
      <div className="user-card-header">
        <div className="user-info">
          <span className="user-label">Usuario:</span>
          <span className="user-card-email">{perm.email}</span>
          <span className={`user-level-badge level-${level}`}>
            {level === 3 ? '👑 Admin' : level === 2 ? '👁️ Supervisor' : '👤 Usuario'}
          </span>
        </div>
        {perm.assignedSheet && (
          <div className="assigned-sheet-badge">
            <span className="assigned-sheet-label">HOJA ASIGNADA:</span>
            <span className="assigned-sheet-name">📋 {perm.assignedSheet}</span>
          </div>
        )}
      </div>

      {!isAdminOrSupervisor && (
        <>
          {isAdmin && (
            <button
              className="btn-assign"
              onClick={() => onAssignIds(perm)}
            >
              Asignar IDs Permitidos
            </button>
          )}

          <div className="ids-section">
            {perm.allowedIds.length === 0 ? (
              <span className="no-ids">Sin IDs asignados</span>
            ) : (
              <>
                <button
                  type="button"
                  className="btn-ver-ids"
                  onClick={() => onToggleExpand(originalIdx)}
                >
                  {expandedIdsIndex === originalIdx ? '▼ Ocultar IDs' : `▶ Ver IDs (${perm.allowedIds.length})`}
                </button>
                {expandedIdsIndex === originalIdx && (
                  <>
                    <div className="ids-tags">
                      {perm.allowedIds.map((id, i) => (
                        <span key={i} className="id-tag">{id}</span>
                      ))}
                    </div>
                    <span className="ids-hint">
                      Este usuario solo verá las filas con estos IDs en la primera columna
                    </span>
                  </>
                )}
              </>
            )}
          </div>
        </>
      )}
    </div>
  )
}

export const UserCard = memo(UserCardInner)
