#!/usr/bin/env python3
"""Extract plate and solid stress from an OP2 file.

Demonstrates:
- CQUAD4/CTRIA3 plate stress extraction (BILINEAR and CENTROID)
- CHEXA/CPENTA/CTETRA solid stress extraction
- Finding maximum von Mises stress across all elements

Usage:
    python extract_op2_stresses.py model.op2
"""
import sys
import numpy as np
from pyNastran.op2.op2 import OP2


def extract_plate_stress(op2_obj: OP2) -> None:
    """Extract plate (shell) element stresses."""
    print("=" * 60)
    print("PLATE STRESS")
    print("=" * 60)

    # Stress column indices:
    # 0: fiber_distance, 1: oxx, 2: oyy, 3: txy,
    # 4: angle, 5: omax, 6: omin, 7: von_mises

    for label, stress_dict in [
        ('CQUAD4', op2_obj.cquad4_stress),
        ('CTRIA3', op2_obj.ctria3_stress),
    ]:
        if not stress_dict:
            continue

        for subcase_id, stress in sorted(stress_dict.items()):
            data = stress.data  # (ntimes, n, 8)
            elem_node = stress.element_node  # (n, 2)

            # Centroid-only results (node_id == 0 for centroid rows)
            centroid_mask = elem_node[:, 1] == 0
            centroid_eids = elem_node[centroid_mask, 0]
            n_elements = len(centroid_eids)

            print(f"\n{label} — Subcase {subcase_id}: "
                  f"{n_elements} elements")

            for itime in range(data.shape[0]):
                vm_centroid = data[itime, centroid_mask, 7]

                idx_max = np.argmax(vm_centroid)
                max_eid = centroid_eids[idx_max]
                max_vm = vm_centroid[idx_max]

                print(f"\n  Time step {itime}:")
                print(f"    Max von Mises (centroid) = {max_vm:.3e} "
                      f"at element {max_eid}")
                print(f"    Mean von Mises           = "
                      f"{np.mean(vm_centroid):.3e}")

                # Top 5 stressed elements
                top5_idx = np.argsort(vm_centroid)[-5:][::-1]
                print(f"\n    Top 5 stressed {label} elements:")
                print(f"    {'EID':>8s}  {'oxx':>12s}  {'oyy':>12s}  "
                      f"{'txy':>12s}  {'vonMises':>12s}")
                for idx in top5_idx:
                    row = np.where(centroid_mask)[0][idx]
                    oxx = data[itime, row, 1]
                    oyy = data[itime, row, 2]
                    txy = data[itime, row, 3]
                    vm = data[itime, row, 7]
                    eid = centroid_eids[idx]
                    print(f"    {eid:8d}  {oxx:12.3e}  {oyy:12.3e}  "
                          f"{txy:12.3e}  {vm:12.3e}")

            # Also show corner-node results if BILINEAR
            if not centroid_mask.all():
                print(f"\n    Note: BILINEAR output detected — "
                      f"{data.shape[1]} total rows "
                      f"({n_elements} centroids + corner nodes)")


def extract_solid_stress(op2_obj: OP2) -> None:
    """Extract solid element stresses."""
    print("\n" + "=" * 60)
    print("SOLID STRESS")
    print("=" * 60)

    # Solid stress columns:
    # 0: oxx, 1: oyy, 2: ozz, 3: txy, 4: tyz, 5: txz,
    # 6: omax, 7: omid, 8: omin, 9: von_mises

    for label, stress_dict in [
        ('CHEXA', op2_obj.chexa_stress),
        ('CPENTA', op2_obj.cpenta_stress),
        ('CTETRA', op2_obj.ctetra_stress),
    ]:
        if not stress_dict:
            continue

        for subcase_id, stress in sorted(stress_dict.items()):
            data = stress.data  # (ntimes, n, 10)
            elem_node = stress.element_node  # (n, 2)

            centroid_mask = elem_node[:, 1] == 0
            centroid_eids = elem_node[centroid_mask, 0]
            n_elements = len(centroid_eids)

            print(f"\n{label} — Subcase {subcase_id}: "
                  f"{n_elements} elements")

            for itime in range(data.shape[0]):
                vm_centroid = data[itime, centroid_mask, 9]

                idx_max = np.argmax(vm_centroid)
                max_eid = centroid_eids[idx_max]
                max_vm = vm_centroid[idx_max]

                print(f"\n  Time step {itime}:")
                print(f"    Max von Mises (centroid) = {max_vm:.3e} "
                      f"at element {max_eid}")

                # Top 5
                top5_idx = np.argsort(vm_centroid)[-5:][::-1]
                print(f"\n    Top 5 stressed {label} elements:")
                print(f"    {'EID':>8s}  {'oxx':>12s}  {'oyy':>12s}  "
                      f"{'ozz':>12s}  {'vonMises':>12s}")
                for idx in top5_idx:
                    row = np.where(centroid_mask)[0][idx]
                    oxx = data[itime, row, 0]
                    oyy = data[itime, row, 1]
                    ozz = data[itime, row, 2]
                    vm = data[itime, row, 9]
                    eid = centroid_eids[idx]
                    print(f"    {eid:8d}  {oxx:12.3e}  {oyy:12.3e}  "
                          f"{ozz:12.3e}  {vm:12.3e}")


def find_global_max_vm(op2_obj: OP2) -> None:
    """Find the global maximum von Mises stress across all element types."""
    print("\n" + "=" * 60)
    print("GLOBAL MAX VON MISES")
    print("=" * 60)

    max_vm = 0.0
    max_info = ""

    # Check all element types
    for label, stress_dict, vm_col in [
        ('CQUAD4', op2_obj.cquad4_stress, 7),
        ('CTRIA3', op2_obj.ctria3_stress, 7),
        ('CHEXA', op2_obj.chexa_stress, 9),
        ('CPENTA', op2_obj.cpenta_stress, 9),
        ('CTETRA', op2_obj.ctetra_stress, 9),
    ]:
        if not stress_dict:
            continue

        for subcase_id, stress in stress_dict.items():
            elem_node = stress.element_node
            centroid_mask = elem_node[:, 1] == 0
            centroid_eids = elem_node[centroid_mask, 0]

            for itime in range(stress.data.shape[0]):
                vm = stress.data[itime, centroid_mask, vm_col]
                idx = np.argmax(vm)
                if vm[idx] > max_vm:
                    max_vm = vm[idx]
                    max_info = (f"{label} EID={centroid_eids[idx]} "
                                f"SC={subcase_id} t={itime}")

    if max_info:
        print(f"\n  Global max von Mises = {max_vm:.3e}")
        print(f"  Location: {max_info}")
    else:
        print("\n  No stress results found")


def main(op2_filename: str) -> None:
    op2 = OP2()
    op2.read_op2(op2_filename)

    extract_plate_stress(op2)
    extract_solid_stress(op2)
    find_global_max_vm(op2)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python extract_op2_stresses.py <model.op2>")
        sys.exit(1)
    main(sys.argv[1])
