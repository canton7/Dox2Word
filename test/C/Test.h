/**
 * @defgroup Test Test Group
 * @ingroup Outer
 * 
 * Brief description.
 * 
 * Detailed description.
 * 
 * <strong>Bold</strong> <s>strike</s> <em>italic @c typewriter text</em>
 * 
 * @warning Warning with <strong>bold text</strong>
 *
 * A numbered list:
 *  1. Item 1
 *    1. Item 1.1
 *  2. Item 2
 * 
 * A bulleted list:
 * - Item 1
 *   - Item 1.1
 * - Item2
 * 
 * @code{.c}
 * This is a code sample
 * With multiple lines
 * @endcode
 * 
 * @todo This is a TODO
 * 
 * Table with alignment:
 * 
 * | Right | Center | Left  |
 * | ----: | :----: | :---- |
 * | 10    | 10     | 10    |
 * | This is some long text | This is some long text   | This is some long text  |
 * | ^     | Merge  ||
 * 
 * Dot diagram:
 * @dot
 * digraph example {
 *     node [shape=record, fontname=Helvetica, fontsize=10];
 *     b [ label="class B" ];
 *     c [ label="class C" ];
 *     b -> c [ arrowhead="open", style="dashed" ];
 * }
 * @enddot
 * 
 * Dot from external file:
 * @dotfile Test.dot "Dot from external file"
 * 
 * Some symbols: © -- ±<br>
 * A newline with a link: https://www.google.com
 * 
 * @{
 * @file
 */

#include "ReferencedFile.h"

#include <stdint.h>
#include <stdbool.h>

/// Constant macro
#define CONSTANT_MACRO (uint32_t)3

/// Single-line function-like macro
///
/// @param x Input parameter
/// @returns Return value
/// @retval 1 Retval 1
/// @retval 2 Retval 2
#define FUNCTION_LIKE_MACRO(x) (uint32_t)(x * CONSTANT_MACRO)

/// Multi-line function-like macro
#define MULTI_FUNCTION_LIKE_MACRO(x, y) \
    do { \
        printf("%i\n", x + y); \
    } while (0)

/**
 * An enum
 */
typedef enum Enum_tag
{
    Enum_One,       ///< Implicit value
    Enum_Two = 2,   ///< Explicit value

    /// Brief description.
    /// Detailed description.
    Enum_Three,

    /// Explicit value which is another member
    Enum_Four = Enum_Three
} Enum_t;

/**
 * A structure
 * 
 * Detailed description
 */
typedef struct Struct_tag
{
    /// Member 'i'
    ///
    /// Detailed description paragraph 1
    ///
    /// Detailed description paragraph 2
	int32_t i;

    Enum_t j; ///< Single-line comment
} Struct_t;

/**
 * A test union
 */
typedef union Union_tag
{
    /// Union member 1
    uint8_t one;

    /// Union member 2
    Enum_t two;
} Union_t;

/**
 * A struct containing an anonymous union
 */
typedef struct StructWithUnion_tag
{
    /// The type discriminator
    uint8_t type;
    union
    {
        uint16_t one : 4; ///< A bitfield
        uint32_t two;   ///< A normal field
    };
} StructWithUnion_t;

/**
 * A typedef to another, linked type
 */
typedef StructWithUnion_t TypedefType_t;

/**
 * Global extern variable
 */
extern Struct_t GlobalVariable;

/**
 * Global variable with multi-line initializer
 */
static const int32_t MultilineGlobalVariable = 1 +
    (2 * 3) + CONSTANT_MACRO;

/**
 * A function which takes void and returns void
 */
void VoidFunction(void);

/**
 * A function which has an empty parameter list
 */
void VoidlessFunction();

/**
 * A function with parameters.
 * 
 * Detailed description.
 * 
 * @param[in] a Parameter a
 * @param[out] b Parameter b
 * @param[in,out] c Parameter C
 * @param d Parameter d with @c typewriter text
 * @returns Return value
 * @retval 1 Retval 1
 * @retval 2 Retval 2
 * 
 * Text after parameter docs
 */
TypedefType_t FunctionWithParameters(uint8_t *a, const TypedefType_t* b, Struct_t* const c, Enum_t d);

/**
 * A static inline function
 * 
 * @param x input
 * @returns @c x doubled
 */
static inline int32_t StaticFunction(int32_t x)
{
    return x * 2;
}

/// @}
