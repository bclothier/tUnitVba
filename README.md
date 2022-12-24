# tUnitVba Proof of concept
A proof of concept demonstrating twInBasic solution running as a VBE add-in to run unit tests. This demonstrates the ability to:

Built with twinBASIC IDE version 178

* Work with memory directly
* Obtain ITypeInfos from VBA project which in turn enables access to additional features not exposed via VBIDE addin
* Allows host-agnostic execution of code
* Supports the special `Print` syntax to format the output

This was frankenstein'd together with various code contributions from the following sources & people:

* [GreedQuest's Gist](https://gist.github.com/Greedquest/faa9b2cd39a2503e84dc297c3d961f73) (Note: this is a VBx port of C# implementation as originally used in [Rubberduck-VBA](https://github.com/rubberduck-vba/Rubberduck))
* [Cristian Buse's Memory Tools](https://github.com/cristianbuse/VBA-MemoryTools) as used in GreedQuest's original gist.
* [Mike Wolfe's DocTest implementation](https://nolongerset.com/python-inspired-doc-tests-in-vba/)
* [Eduardo's example of IVBprint implementation](https://www.vbforums.com/showthread.php?891891-(VB6)-Implement-the-Print-method-on-any-object)

Here is a very rough and simple demonstration:

<video src='https://user-images.githubusercontent.com/2367644/209445420-3dcb89a0-6dcf-45f7-bb51-fcc14fb82c9a.mp4' width=180/>

## Known Limitations & Bugs

1. The `Print` implementation is not 100% equivalent as the `Debug.Print` in all cases; one edge case involves switching between the `,` and `;` in the same `Print` statement which is not easily replicated. 
2. Mike's original code used `Eval()` which is Access-specific. That means it could at least evaluate code like `foo(1+2)`. In this implementation, it cannot since the 1+2` would be passed as a string rather than an expression. Therefore, it is not possible to use expressions in the DocTests. 
3. There is no compile-time validation of the DocTests since they are comments. For complex testing requirements, use [Rubberduck VBA addin](https://github.com/rubberduck-vba/Rubberduck) instead. 
4. The `ITypeLib` and `ITypeInfo` are not fully implemented & tested. 
5. The doc tests are "discovered" via regex parsing with the expectation that the line will be a valid call statement to a single procedure. 
6. Private access is not working yet. This requires using the `ITypeInfo::Funcs` which is not implemented fully. 
