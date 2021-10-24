/**
 * @defgroup CommonUnit Common - Unit
 * @ingroup Common
 * 
 * Common Unit is a Unit in the Common module.
 * 
 * <strong>Bold and <em>italic and @c typewriter </em></strong>
 * 
 * @warning This is a very important warning!
 *
 * A list of things:
 *  1. Thing 1
 *    1. Thing 1.1
 *  2. Thing 2
 * 
 * @{
 * @file
 */

#include <stdint.h>
#include <stdbool.h>

/// This is a macro
///
/// @param x Thing
/// @returns @c x doubled
#define COMMON_UNIT_MACRO(x) (uint32_t)(x * 2)

/// This is a constant
#define COMMON_UNIT_CONST (uint32_t)3

/**
 * This is a global variable
 */
extern bool CommonUnit_Global;

/**
 * A test structure
 * 
 * More description of the test structure
 */
typedef struct CommonUnit_Struct_tag
{
    /// Member 'i'
    ///
    /// This member is so important it gets a detailed description.
    ///
    /// With two paragraphs in it.
	int32_t i;

    bool j; ///< Just 'j'
} CommonUnit_Struct_t;

/**
 * This is a custom typedef
 */
typedef uint8_t CommonUnit_TypedefType_t;

/**
 * This is the common unit's test file
 * 
 * This is a longer description of the function.
 * 
 * And another paragraph of description.
 * 
 * @param p Some parameters
 * @returns A value
 * 
 * Some more text afterwards.
 */
void CommonUnit_Test(const uint8_t p);

/**
 * Another function in CommonUnit
 * 
 * Usage:
 * @code{.c}
 * int result = CommonUnit_Another();
 * Foo x = rand();
 * @endcode
 *
 * @param in In documentation @c yay
 * @param foo Foo documentation
 */
CommonUnit_TypedefType_t CommonUnit_Another(const CommonUnit_Struct_t* in, uint8_t* const foo);

/**
 * Thingy docs
 */
void Thingy(void);

/// @}
