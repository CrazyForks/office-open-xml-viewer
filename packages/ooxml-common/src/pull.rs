//! Stable, format-neutral wire envelopes used by bounded pull cursors.

const INSUFFICIENT_CREDIT_PREFIX: &str = "OOXML_INSUFFICIENT_CREDIT:";

/// Report that one indivisible pull unit needs more credit. The exact JSON
/// envelope is decoded by the shared TypeScript worker boundary; format code
/// must not parse or depend on the human-readable message.
pub fn insufficient_credit_error(required_bytes: usize, offered_bytes: usize) -> String {
    format!(
        "{INSUFFICIENT_CREDIT_PREFIX}{{\"code\":\"ooxml-insufficient-credit\",\"requiredBytes\":{required_bytes},\"offeredBytes\":{offered_bytes}}}"
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn emits_the_shared_machine_envelope() {
        assert_eq!(
            insufficient_credit_error(2048, 1024),
            "OOXML_INSUFFICIENT_CREDIT:{\"code\":\"ooxml-insufficient-credit\",\"requiredBytes\":2048,\"offeredBytes\":1024}"
        );
    }
}
