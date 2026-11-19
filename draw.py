import random


class DrawConfigurationError(ValueError):
    """The current roster and saved assignments cannot support a complete draw."""


def has_recipient(assignment):
    partner = assignment.get("partner") or ""
    return bool(partner) and partner != "SIN PAREJA" and not partner.startswith("SIN PAREJA (")


def plan_draw(participants, assignments, *, shuffle=False):
    """Match every remaining giver to a unique recipient from a different family."""
    missing_groups = sum(not group for group in participants.values())
    if missing_groups:
        raise DrawConfigurationError(
            f"Completa «Grupo familiar» en el Excel de participantes. Faltan {missing_groups} registros."
        )

    assigned_givers = set()
    assigned_recipients = set()
    for assignment in assignments:
        if not has_recipient(assignment):
            continue
        giver, recipient = assignment["name"], assignment["partner"]
        if giver not in participants or recipient not in participants:
            raise DrawConfigurationError(
                f"Revisa la asignación de {giver}. Alguno de sus nombres falta en el Excel de participantes."
            )
        if participants[giver] == participants[recipient]:
            raise DrawConfigurationError(
                f"La asignación de {giver} está dentro de su grupo familiar. Corrígela desde el panel."
            )
        if giver in assigned_givers or recipient in assigned_recipients:
            raise DrawConfigurationError("Hay nombres repetidos en las asignaciones guardadas. Revísalas desde el panel.")
        assigned_givers.add(giver)
        assigned_recipients.add(recipient)

    givers = [name for name in participants if name not in assigned_givers]
    recipients = [name for name in participants if name not in assigned_recipients]
    candidates = {
        giver: [recipient for recipient in recipients if participants[giver] != participants[recipient]]
        for giver in givers
    }
    if shuffle:
        random.shuffle(givers)
        for choices in candidates.values():
            random.shuffle(choices)

    matched_givers = {}

    def find_recipient(giver, visited):
        for recipient in candidates[giver]:
            if recipient in visited:
                continue
            visited.add(recipient)
            previous_giver = matched_givers.get(recipient)
            if previous_giver is None or find_recipient(previous_giver, visited):
                matched_givers[recipient] = giver
                return True
        return False

    for giver in givers:
        if not find_recipient(giver, set()):
            raise DrawConfigurationError(
                "Con estos grupos familiares y las asignaciones guardadas, "
                "no es posible encontrar una pareja para todos. Revisa el Excel y el panel."
            )
    return {giver: recipient for recipient, giver in matched_givers.items()}
