/**
 * @addtogroup Test
 * @{
 * @file
 */

int32_t GlobalVariable = 3;

/**
 * This is a static variable with a value (documented for some reason)
 */
static uint32_t StaticWithValue = 3;

uint8_t FunctionWithParameters(uint8_t* a, const TypedefType_t* b, Struct_t* const c, Enum_t d)
{
    ReferencedFunction();
    VoidFunction();
    return 0;
}

static void StaticFunction(void)
{
}

/// @}
