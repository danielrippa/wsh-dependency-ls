
  fail = (message, errorlevel, error) !->

    write = -> WScript.StdErr.Write [ (arg) for arg in arguments ] * ' '
    writeln = -> write '\n' ; write ...

    write-kv = (k, v) -> writeln "#k: #v"

    write message

    for k,v of error

      switch k

        | 'number' =>

          # https://web.archive.org/web/20240709175529/https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-hresult_facility

          # extract facility code and error code from HRESULT

          write-kv \facility, (v .>>. 16) .&. 0x1fff
          write-kv \number, v .&. 0xffff

        else write-kv k, v

    WScript.Quit errorlevel

  dependency = do ->

    WScript.Arguments.Unnamed

      argc = ..Count
      argv = [ (..Item index) for index til argc ]

    #

    new ActiveXObject 'WScript.Shell'

      current-folder = ..CurrentDirectory

    #

    new ActiveXObject 'Scripting.FileSystemObject'

      file-exists = -> ..FileExists it
      folder-exists = -> ..FolderExists it

      open-file = -> ..OpenTextFile it, 1

    #

    get-content = -> (open-file it) => content = ..ReadAll! ; ..Close! ; return content

    #

    lcase = (.to-lower-case!)

    #

    slash = '\\'

    split-path = (/ "#slash")
    join-path  = (* "#slash")

    get-path = -> it |> split-path |> (.slice 0, -1) |> join-path

    #

    string-as-array = do ->

      us = String.from-char-code 31

      replace-crlf = (.replace /\r\n/g, us)
      replace-lf   = (.replace /\n/g, us)

      string-as-units = -> it |> replace-crlf |> replace-lf

      units-as-array = (.split us)

      #

      -> it |> string-as-units |> units-as-array

    #

    trim = do ->

      trim-regex = /^\s+|\s+$/g

      (.replace trim-regex, '')

    #

    read-configuration-file = (filepath) ->

      configuration = {}

      if file-exists filepath

        configuration-lines = filepath |> get-content |> string-as-array

        for line, line-number in configuration-lines

          line = trim line

          if line is ''
            continue

          if (line.char-at 0) is '#'
            continue

          space-index = line.index-of ' '

          throw new Error "Invalid configuration file syntax at line (#line-number) '#line' of configuration file '#filename'" \
            if space-index is -1

          key = line.slice 0, space-index

          value = line.slice space-index + 1

          configuration[ key ] = value

      configuration

    #

    namespace-path-manager = do ->

      script-path = get-path argv.0

      configuration-filename = 'namespaces.conf'

      configuration-filepath = [ script-path, configuration-filename ] |> join-path

      if not file-exists configuration-filepath

        configuration-filepath = [ current-folder, configuration-filename ] |> join-path

      configuration-namespaces = read-configuration-file configuration-filepath

      namespaces = '': current-folder

      get-qualified-namespace-path = (qualified-namespace) ->

        # registered namespaces

        namespace-path = namespaces[ qualified-namespace ]

        if namespace-path isnt void

          return namespace-path

        # configuration-namespaces

        namespace-path = configuration-namespaces[ qualified-namespace ]

        if namespace-path isnt void

          if folder-exists namespace-path

            namespaces[ qualified-namespace ] := namespace-path
            return namespace-path

          throw new Error "Folder '#namespace-path' for namespace '#qualified-namespace' in configuration file '#configuration-filename' not found."

        # script-path

        qualified-namespace-path = qualified-namespace |> (/ '.') |> join-path

        namespace-path = [ script-path, qualified-namespace-path ] |> join-path

        if folder-exists namespace-path

          namespaces[ qualified-namespace ] := namespace-path
          return namespace-path

        # current-folder path

        namespace-path = [ current-folder, qualified-namespace-path ] |> join-path

        if folder-exists namespace-path

          namespaces[ qualified-namespace ] := namespace-path
          return namespace-path

        throw new Error "Folder for namespace '#qualified-namespace' not found."

      {
        get-qualified-namespace-path
      }

    #

    parse-qualified-dependency-name = (qualified-dependency-name) ->

      [ ...namespaces, dependency-name ] = qualified-dependency-name / '.'

      qualified-namespace = namespaces * '.' |> lcase

      { qualified-namespace, dependency-name }

    #

    dependency-builder = do ->

      build-dependency = (qualified-dependency-name) ->

        { qualified-namespace, dependency-name } = parse-qualified-dependency-name qualified-dependency-name

        filename = [ dependency-name, 'js' ] * '.'

        namespace-path = namespace-path-manager.get-qualified-namespace-path qualified-namespace

        dependency-full-path = [ namespace-path, filename ] |> join-path

        if not file-exists dependency-full-path

          throw new Error "Dependency file '#dependency-full-path' not found."

        try eval get-content dependency-full-path
        catch => fail "Unable to load dependency '#qualified-dependency-name' (#dependency-full-path)", 2, e

      {
        build-dependency
      }

    #

    dependency-manager = do ->

      dependencies = {}

      get-dependency = (qualified-dependency-name) ->

        qname = lcase qualified-dependency-name

        result = dependencies[ qname ]

        if result is void

          result = dependency-builder.build-dependency qualified-dependency-name

          dependencies[ qname ] := result

        result

      {
        get-dependency
      }


    #

    dependency = (qualified-dependency-name) ->

      dependency-manager.get-dependency qualified-dependency-name

    if argc > 0

      script-path = argv.0

      failure = null

      try script-source = get-content script-path
      catch => failure = message: "Unable to read script '#script-path'", error: e

      unless failure?

        try eval script-source
        catch => failure = message: "Unable to execute script '#script-path'", error: e

      if failure? then failure => fail ..message, 1, ..error

    dependency
